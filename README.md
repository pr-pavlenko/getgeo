# GetGeo - Google Sheets Custom Function for Geocoding

Use the formula as before: =getGeo("O", "P" & ROW())

This repository provides a custom function for Google Sheets that geocodes a city and state into latitude and longitude using the geocode.maps.co API. It allows users to input the column letters for the City and State, and the function will return the coordinates as latitude,longitude.

Features

	•	Geocode city and state from specified columns in your Google Sheets.
	•	Supports dynamic row numbers using the ROW() function.
	•	Automatically handles API rate limiting (1 request per second) with a delay.
	•	Uses the free Geocoding API from geocode.maps.co (1 Request/Second 5,000 /day)

Usage

After adding the script, you can use the getGeo function directly within your Google Sheets.

Formula Syntax

=getGeo("CityColumn", "StateColumn" & ROW())

Example

Assume the following:
	•	City is in column O.
	•	State is in column P.

You can use the formula like this:
=getGeo("O", "P" & ROW())

This will return the latitude and longitude of the city and state in columns O and P for the corresponding row.

Rate Limiting

The geocode.maps.co API has a 1 request per second limit for the free plan. To ensure that the script works correctly without exceeding the rate limit, a delay is introduced between each request using Utilities.sleep(1000) (1 second delay).




