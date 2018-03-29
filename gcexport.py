#!/usr/bin/python

"""
File: gcexport.py
Author: Kyle Krafka (https://github.com/kjkjava/)
Date: April 28, 2015

Description:	Use this script to export your fitness data from Garmin Connect.
				See README.md for more information.
"""

from urllib.parse import urlencode
from datetime import datetime
from getpass import getpass
from sys import argv
from os.path import isdir
from os.path import isfile
from os import mkdir
from os import remove
from xml.dom.minidom import parseString

import urllib.request, urllib.error, urllib.parse, http.cookiejar, json
from fileinput import filename

import argparse
import zipfile
import traceback

# url is a string, post is a dictionary of POST parameters, headers is a dictionary of headers.
def _http_request(opener, url, post=None, headers={}):
	request = urllib.request.Request(url)
	request.add_header('User-Agent', 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/1337 Safari/537.36')  # Tell Garmin we're some supported browser.
	for header_key, header_value in headers.items():
		request.add_header(header_key, header_value)
	if post:
		post = urlencode(post).encode('utf-8')  # Convert dictionary to POST parameter string.
	response = opener.open(request, data=post)  # This line may throw a urllib2.HTTPError.

	# N.B. urllib2 will follow any 302 redirects. Also, the "open" call above may throw a urllib2.HTTPError which is checked for below.
	if response.getcode() != 200:
		raise Exception('Bad return code (' + response.getcode() + ') for: ' + url)

	return response.read()

class GarminConnect(object):
	# URLs for various services.
	LOGIN_URL     	= ('https://sso.garmin.com/sso/login?service=https%3A%2F%2Fconnect.garmin.com'
		'%2Fpost-auth%2Flogin&webhost=olaxpw-connect04&source=https%3A%2F%2Fconnect.garmin.com%2Fen-US%2Fsignin&redirectAfterAccount'
		'LoginUrl=https%3A%2F%2Fconnect.garmin.com%2Fpost-auth%2Flogin&redirectAfterAccountCreationUrl=https%3A%2F%2Fconnect.garmin.com'
		'%2Fpost-auth%2Flogin&gauthHost=https%3A%2F%2Fsso.garmin.com%2Fsso&locale=en_US&id=gauth-widget&cssUrl=https%3A%2F%2Fstatic.garmincdn.com'
		'%2Fcom.garmin.connect%2Fui%2Fcss%2Fgauth-custom-v1.1-min.css&clientId=GarminConnect&rememberMeShown=true&rememberMeChecked=false'
		'&createAccountShown=true&openCreateAccount=false&usernameShown=false&displayNameShown=false&consumeServiceTicket=false&initialFocus=true&embedWidget=false&generateExtraServiceTicket=false')
	POST_AUTH_URL 	= 'https://connect.garmin.com/post-auth/login?'
	SEARCH_URL    	= 'http://connect.garmin.com/proxy/activity-search-service-1.2/json/activities?'
	GPX_ACTIVITY_URL = 'https://connect.garmin.com/modern/proxy/download-service/export/gpx/activity/'
	TCX_ACTIVITY_URL = 'https://connect.garmin.com/modern/proxy/download-service/export/tcx/activity/'
	ORIGINAL_ACTIVITY_URL = 'http://connect.garmin.com/proxy/download-service/files/activity/'

	def __init__(self, username, password):
		cookie_jar = http.cookiejar.CookieJar()
		self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))

		# Maximum number of activities you can request at once.  Set and enforced by Garmin.
		limit_maximum = 100

		# Initially, we need to get a valid session cookie, so we pull the login page.
		_http_request(self.opener, self.LOGIN_URL)

		# Now we'll actually login.
		post_data = {'username': username, 'password': password, 'embed': 'true', 'lt': 'e1s1', '_eventId': 'submit', 'displayNameRequired': 'false'}  # Fields that are passed in a typical Garmin login.
		_http_request(self.opener, self.LOGIN_URL, post_data)

		# Get the key.
		# TODO: Can we do this without iterating?
		login_ticket = None
		for cookie in cookie_jar:
			if cookie.name == 'CASTGC':
				login_ticket = cookie.value
				break

		if not login_ticket:
			raise Exception('Did not get a ticket cookie. Cannot log in. Did you enter the correct username and password?')

		# Chop of 'TGT-' off the beginning, prepend 'ST-0'.
		login_ticket = 'ST-0' + login_ticket[4:]

		_http_request(self.opener, self.POST_AUTH_URL + 'ticket=' + login_ticket)

		# We should be logged in now.

	def download(self, directory, fileFormat, count, unzip):		
		# Create directory for data files.
		if isdir(directory) and count != 'new':
			print('Warning: Output directory already exists. Will skip already-downloaded files.')
		elif not isdir(directory) and count == 'new':
			raise Exception('Error: Directory does not exist.')
		
		if not isdir(directory):
			mkdir(directory)

		download_all = False
		if count == 'all' or count == 'new':
			# If the user wants to download all activities, first download one,
			# then the result of that request will tell us how many are available
			# so we will modify the variables then.
			total_to_download = 1
			download_all = True
		else:
			total_to_download = int(count)
		total_downloaded = 0

		# This while loop will download data from the server in multiple chunks, if necessary.
		while total_downloaded < total_to_download:
			# Maximum of 100... 400 return status if over 100.  So download 100 or whatever remains if less than 100.
			if total_to_download - total_downloaded > 100:
				num_to_download = 100
			else:
				num_to_download = total_to_download - total_downloaded

			search_params = {'start': total_downloaded, 'limit': num_to_download}
			# Query Garmin Connect
			result = _http_request(self.opener, self.SEARCH_URL + urlencode(search_params))
			json_results = json.loads(result.decode('utf-8'))  # TODO: Catch possible exceptions here.
			
			if download_all:
				# Modify total_to_download based on how many activities the server reports.
				total_to_download = int(json_results['results']['totalFound'])
				# Do it only once.
				download_all = False

			# Pull out just the list of activities.
			activities = json_results['results']['activities']

			# Process each activity.
			for a in activities:
				# Display which entry we're working on.
				activity = a['activity']
				activityId = activity['activityId']
				print('Garmin Connect activity: [' + str(activityId) + ']', end=' ')
				print(a['activity']['activityName'])

				print('\t', end='')
				if 'activitySummary' in a['activity']:
					activity_summary = activity['activitySummary']
					if 'BeginTimestamp' in activity_summary:
						print(activity_summary['BeginTimestamp']['display'] + ',', end=' ')
					else:
						print('??:??:??,', end=' ')
					if 'SumElapsedDuration' in activity_summary:
						print(activity_summary['SumElapsedDuration']['display'] + ',', end=' ')
					else:
						print('??:??:??,', end=' ')
					if 'SumDistance' in activity_summary:
						print(a['activity']['SumDistance']['withUnit'])
					else:
						print('0.00 Miles')
				else:
					print('No summary.')

				if fileFormat == 'gpx':
					data_filename = directory + '/activity_' + str(activityId) + '.gpx'
					download_url = self.GPX_ACTIVITY_URL + str(activityId) + '?full=true'
				elif fileFormat == 'tcx':
					data_filename = directory + '/activity_' + str(activityId) + '.tcx'
					download_url = self.TCX_ACTIVITY_URL + str(activityId) + '?full=true'
				elif fileFormat == 'original':
					data_filename = directory + '/activity_' + str(activityId) + '.zip'
					fit_filename = directory + '/' + str(activityId) + '.fit'
					download_url = self.ORIGINAL_ACTIVITY_URL + str(activityId)
				else:
					raise Exception('Unrecognized file format.')

				if isfile(data_filename):
					print('\tData file already exists; skipping...')
					if count == 'new':
						return
					else:
						continue
				if fileFormat == 'original' and isfile(fit_filename):  # Regardless of unzip setting, don't redownload if the ZIP or FIT file exists.
					print('\tFIT data file already exists; skipping...')
					if count == 'new':
						return
					else:
						continue

				# Download the data file from Garmin Connect.
				# If the download fails (e.g., due to timeout), this script will die, but nothing
				# will have been written to disk about this activity, so just running it again
				# should pick up where it left off.
				print('\tDownloading file...', end=' ')

				try:
					data = _http_request(self.opener, download_url)
				except urllib.error.HTTPError as e:
					# Handle expected (though unfortunate) error codes; die on unexpected ones.
					if e.code == 500 and fileFormat == 'tcx':
						# Garmin will give an internal server error (HTTP 500) when downloading TCX files if the original was a manual GPX upload.
						# Writing an empty file prevents this file from being redownloaded, similar to the way GPX files are saved even when there are no tracks.
						# One could be generated here, but that's a bit much. Use the GPX format if you want actual data in every file,
						# as I believe Garmin provides a GPX file for every activity.
						print('Writing empty file since Garmin did not generate a TCX file for this activity...', end=' ')
						data = b''
					elif e.code == 404 and fileFormat == 'original':
						# For manual activities (i.e., entered in online without a file upload), there is no original file.
						# Write an empty file to prevent redownloading it.
						print('Writing empty file since there was no original activity data...', end=' ')
						data = b''
					else:
						raise Exception('Failed. Got an unexpected HTTP error (' + str(e.code) + ').')

				save_file = open(data_filename, 'wb')
				save_file.write(data)
				save_file.close()

				if fileFormat == 'gpx':
					# Validate GPX data. If we have an activity without GPS data (e.g., running on a treadmill),
					# Garmin Connect still kicks out a GPX, but there is only activity information, no GPS data.
					# N.B. You can omit the XML parse (and the associated log messages) to speed things up.
					gpx = parseString(data)
					gpx_data_exists = len(gpx.getElementsByTagName('trkpt')) > 0

					if gpx_data_exists:
						print('Done. GPX data saved.')
					else:
						print('Done. No track points found.')
				elif fileFormat == 'original':
					if len(data) > 0:
						if unzip and data_filename[-3:].lower() == 'zip':  # Even manual upload of a GPX file is zipped, but we'll validate the extension.
							print("Unzipping and removing original files...", end=' ')
							zip_file = open(data_filename, 'rb')
							z = zipfile.ZipFile(zip_file)
							for name in z.namelist():
								z.extract(name, directory)
							zip_file.close()
							remove(data_filename)
					print('Done.')
				else:
					# TODO: Consider validating other formats.
					print('Done.')
			total_downloaded += num_to_download
		# End while loop for multiple chunks.

script_version = '1.0.0'
current_date = datetime.now().strftime('%Y-%m-%d')

parser = argparse.ArgumentParser()

# TODO: Implement verbose and/or quiet options.
# parser.add_argument('-v', '--verbose', help="increase output verbosity", action="store_true")
parser.add_argument('--version', help="print version and exit", action="store_true")
parser.add_argument('--username', help="your Garmin Connect username (otherwise, you will be prompted)", nargs='?')
parser.add_argument('--password', help="your Garmin Connect password (otherwise, you will be prompted)", nargs='?')

parser.add_argument('-c', '--count', nargs='?', default="1",
	help="number of recent activities to download, or 'all', or 'new' (default: 1)")

parser.add_argument('-f', '--format', nargs='?', choices=['gpx', 'tcx', 'original'], default="gpx",
	help="export format; can be 'gpx', 'tcx', or 'original' (default: 'gpx')")

parser.add_argument('-d', '--directory', nargs='?', default='./',
	help="the directory to export to (default: './YYYY-MM-DD_garmin_connect_export')")

parser.add_argument('-u', '--unzip',
	help="if downloading ZIP files (format: 'original'), unzip the file and removes the ZIP file",
	action="store_true")

args = parser.parse_args()

if args.version:
	print(argv[0] + ", version " + script_version)
	exit(0)

print('Welcome to Garmin Connect Exporter!')

username = args.username if args.username else input('Username: ')
password = args.password if args.password else getpass()

try:
	gc = GarminConnect(username, password)
	gc.download(args.directory, args.format, args.count, args.unzip)
except Exception:
	traceback.print_exc()

print('Done!')