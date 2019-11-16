#!/usr/bin/python

"""
File: gcexport.py
Author: Kyle Krafka (https://github.com/kjkjava/)
Date: April 28, 2015

Description:	Use this script to export your fitness data from Garmin Connect.
				See README.md for more information.
"""

from datetime import datetime, timedelta
from getpass import getpass
from xml.dom.minidom import parseString

import urllib, http.cookiejar, json, re
from fileinput import filename

import argparse
import zipfile
import traceback
import logging
import sys, os

# url is a string, post is a dictionary of POST parameters, headers is a dictionary of headers.
def _http_request(opener, url, post=None, headers={}):
	request = urllib.request.Request(url)
	# Tell Garmin we're some supported browser. 
	request.add_header('User-Agent', 'Mozilla/5.0 (Windows NT 6.1; Win64; x64)')
	for header_key, header_value in headers.items():
		request.add_header(header_key, header_value)
	if post:
		post = urllib.parse.urlencode(post).encode('utf-8')  # Convert dictionary to POST parameter string.
	response = opener.open(request, data=post)  # This line may throw a urllib2.HTTPError.

	# N.B. urllib2 will follow any 302 redirects. Also, the "open" call above may throw a urllib2.HTTPError which is checked for below.
	if response.getcode() != 200:
		raise Exception('Bad return code (' + response.getcode() + ') for: ' + url)

	return response.read()

class GarminConnect(object):
	LOGIN_URL		= 'https://sso.garmin.com/sso/login?'
	POST_AUTH_URL 	= 'https://connect.garmin.com/modern/activities?'
	PROFILE_URL 	= "https://connect.garmin.com/modern/profile"
	USERSTATS_URL 	= "https://connect.garmin.com/modern/proxy/userstats-service/statistics/"
	SEARCH_URL    	= 'https://connect.garmin.com/modern/proxy/activitylist-service/activities/search/activities?'
	GPX_ACTIVITY_URL = 'https://connect.garmin.com/modern/proxy/download-service/export/gpx/activity/'
	TCX_ACTIVITY_URL = 'https://connect.garmin.com/modern/proxy/download-service/export/tcx/activity/'
	ORIGINAL_ACTIVITY_URL = 'http://connect.garmin.com/proxy/download-service/files/activity/'

	def __init__(self):
		self.cookie_jar = http.cookiejar.CookieJar()
		self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.cookie_jar))

	def login(self, username, password):
		data = {
			'service': 'https://connect.garmin.com/modern',
			'webhost': 'https://connect.garmin.com',
			'source': 'https://connect.garmin.com/en-US/signin',
			'redirectAfterAccountLoginUrl': 'https://connect.garmin.com/modern',
			'redirectAfterAccountCreationUrl': 'https://connect.garmin.com/modern',
			'gauthHost': 'https://sso.garmin.com/sso',
			'locale': 'en_US',
			'id': 'gauth-widget',
			'cssUrl': 'https://static.garmincdn.com/com.garmin.connect/ui/css/gauth-custom-v1.2-min.css',
			'clientId': 'GarminConnect',
			'rememberMeShown': 'true',
			'rememberMeChecked': 'false',
			'createAccountShown': 'true',
			'openCreateAccount': 'false',
			'displayNameShown': 'false',
			'consumeServiceTicket': 'false',
			'initialFocus': 'true',
			'embedWidget': 'false',
			'generateExtraServiceTicket': 'false',
			'generateNoServiceTicket': 'false',
			'globalOptInShown': 'true',
			'globalOptInChecked': 'false',
			'mobile': 'false',
			'connectLegalTerms': 'true',
			'locationPromptShown': 'true'
			}

		# Initially, we need to get a valid session cookie, so we pull the login page.
		_http_request(self.opener, self.LOGIN_URL + urllib.parse.urlencode(data))

		# Now we'll actually login.
		# Fields that are passed in a typical Garmin login.
		post_data = {
			'username': username,
			'password': password,
			'embed': 'true',
			'lt': 'e1s1',
			'_eventId': 'submit',
			'displayNameRequired': 'false',
			'queryString': urllib.parse.urlencode(data),
			'contextPath': '/sso'
			}
		login_response = _http_request(self.opener, self.LOGIN_URL + urllib.parse.urlencode(data), post_data).decode()

		# extract the ticket from the login response
		pattern = re.compile(r".*\?ticket=([-\w]+)\";.*", re.MULTILINE | re.DOTALL)
		match = pattern.match(login_response)
		if not match:
			raise Exception('Did not get a ticket in the login response. Cannot log in. Did you enter the correct username and password?')
		login_ticket = match.group(1)

		if not login_ticket:
			raise Exception('Did not get a ticket cookie. Cannot log in. Did you enter the correct username and password?')

		_http_request(self.opener, self.POST_AUTH_URL + 'ticket=' + login_ticket)

		# We should be logged in now.

	def download(self, directory, fileFormat, count, unzip):
		# Create directory for data files.		
		if not os.path.isdir(directory):
			os.mkdir(directory)

		if count == 'all' or count == 'new':
			# If the user wants to download all activities, query the userstats
			# on the profile page to know how many are available
			logging.info("Getting display name and user stats via: " + self.PROFILE_URL)
			profile_page = _http_request(self.opener, self.PROFILE_URL).decode('utf-8')
			# write_to_file(args.directory + '/profile.html', profile_page, 'a')

			# extract the display name from the profile page, it should be in there as
			# \"displayName\":\"eschep\"
			pattern = re.compile(r".*\\\"displayName\\\":\\\"([-.\w]+)\\\".*", re.MULTILINE | re.DOTALL)
			match = pattern.match(profile_page)
			if not match:
				raise Exception("Did not find the display name in the profile page.")
			display_name = match.group(1)
			logging.info("Found the display name: " + display_name)

			# Modify total_to_download based on how many activities the server reports.
			user_stats = _http_request(self.opener, self.USERSTATS_URL + display_name).decode('utf-8')
			json_user = json.loads(user_stats)
			total_to_download = int(json_user["userMetrics"][0]["totalActivities"])

			logging.info("Found the activity count: " + str(total_to_download))
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
			result = _http_request(self.opener, self.SEARCH_URL + urllib.parse.urlencode(search_params)).decode('utf-8')
			json_results = json.loads(result)  # TODO: Catch possible exceptions here.
			
			# Pull out just the list of activities.
			activities = json_results

			# Process each activity.
			for activity in activities:
				# Display which entry we're working on.
				activityId = activity['activityId']
				logging.info('Garmin Connect activity: [' + str(activityId) + '] ' + activity['activityName'])

				startTimeLocal = activity['startTimeLocal'] if 'startTimeLocal' in activity else '??:??:??'
				duration = str(timedelta(seconds=activity['duration'])) if 'duration' in activity else '??:??:??'
				distance = str(activity['distance'] / 1000) if 'distance' in activity else '?'
				logging.debug('\t' + 'Start time: ' + startTimeLocal + ', Duration: ' + duration + ', Distance: ' + distance + ' km')

				if fileFormat == 'gpx':
					data_filename = os.path.join(directory, 'activity_' + str(activityId) + '.gpx')
					download_url = self.GPX_ACTIVITY_URL + str(activityId) + '?full=true'
				elif fileFormat == 'tcx':
					data_filename = os.path.join(directory, 'activity_' + str(activityId) + '.tcx')
					download_url = self.TCX_ACTIVITY_URL + str(activityId) + '?full=true'
				elif fileFormat == 'original':
					data_filename = os.path.join(directory, 'activity_' + str(activityId) + '.zip')
					fit_filename = os.path.join(directory, str(activityId) + '.fit')
					download_url = self.ORIGINAL_ACTIVITY_URL + str(activityId)
				else:
					raise Exception('Unrecognized file format.')

				if os.path.isfile(data_filename):
					logging.info('\tData file already exists; skipping...')
					if count == 'new':
						return
					else:
						continue
				if fileFormat == 'original' and os.path.isfile(fit_filename):  # Regardless of unzip setting, don't redownload if the ZIP or FIT file exists.
					logging.info('\tFIT data file already exists; skipping...')
					if count == 'new':
						return
					else:
						continue

				# Download the data file from Garmin Connect.
				# If the download fails (e.g., due to timeout), this script will die, but nothing
				# will have been written to disk about this activity, so just running it again
				# should pick up where it left off.
				logging.debug('\tDownloading file...')

				try:
					data = _http_request(self.opener, download_url)
				except urllib.error.HTTPError as e:
					# Handle expected (though unfortunate) error codes; die on unexpected ones.
					if e.code == 500 and fileFormat == 'tcx':
						# Garmin will give an internal server error (HTTP 500) when downloading TCX files if the original was a manual GPX upload.
						# Writing an empty file prevents this file from being redownloaded, similar to the way GPX files are saved even when there are no tracks.
						# One could be generated here, but that's a bit much. Use the GPX format if you want actual data in every file,
						# as I believe Garmin provides a GPX file for every activity.
						logging.warning('\tWriting empty file since Garmin did not generate a TCX file for this activity...')
						data = b''
					elif e.code == 404 and fileFormat == 'original':
						# For manual activities (i.e., entered in online without a file upload), there is no original file.
						# Write an empty file to prevent redownloading it.
						logging.warning('\tWriting empty file since there was no original activity data...')
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
						logging.info('\tDone. GPX data saved.')
					else:
						logging.info('\tDone. No track points found.')
				elif fileFormat == 'original':
					if len(data) > 0:
						if unzip and data_filename[-3:].lower() == 'zip':  # Even manual upload of a GPX file is zipped, but we'll validate the extension.
							logging.debug("\tUnzipping and removing original files...")
							zip_file = open(data_filename, 'rb')
							z = zipfile.ZipFile(zip_file)
							for name in z.namelist():
								z.extract(name, directory)
							zip_file.close()
							os.remove(data_filename)
					logging.info('\tDone.')
				else:
					# TODO: Consider validating other formats.
					logging.info('\tDone.')
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
	help="the directory to export to (default: './')")

parser.add_argument('-u', '--unzip',
	help="if downloading ZIP files (format: 'original'), unzip the file and removes the ZIP file",
	action="store_true")
	
parser.add_argument('-l', '--log',
	help="enable logging",
	action="store_true")

args = parser.parse_args()

if args.version:
	print(argv[0] + ", version " + script_version)
	exit(0)

print('Welcome to Garmin Connect Exporter!')

logger = logging.getLogger()

# Console logging
logger.setLevel(logging.DEBUG)
handler = logging.StreamHandler()
handler.setLevel(logging.DEBUG)
logger.addHandler(handler)

if (args.log):
	# Error logging to error.log
	handler = logging.FileHandler(os.path.join(args.directory, "error.log"))
	handler.setLevel(logging.ERROR)
	formatter = logging.Formatter("[%(asctime)s] %(message)s")
	handler.setFormatter(formatter)
	logger.addHandler(handler)

	# Debug logging to debug.log
	handler = logging.FileHandler(os.path.join(args.directory, "debug.log"))
	handler.setLevel(logging.DEBUG)
	formatter = logging.Formatter("[%(asctime)s] %(message)s")
	handler.setFormatter(formatter)
	logger.addHandler(handler)

username = args.username if args.username else input('Username: ')
password = args.password if args.password else getpass()

try:
	gc = GarminConnect()
	gc.login(username, password)
	gc.download(args.directory, args.format, args.count, args.unzip)
except Exception as exception:
	logging.error(exception)
	traceback.print_exc()
	sys.exit(-1)

print('Done!')