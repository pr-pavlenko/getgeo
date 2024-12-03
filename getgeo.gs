**
 * Get Geographic Coordinates
 * @Author: Roman Pavlenko
 * @param {string} cityColumn The column letter for the city (e.g., "O").
 * @param {string} stateColumnAndRow The column letter and row for the state (e.g., "P5").
 * @returns {string} Coordinates in the format "latitude,longitude".
 */
function getGeo(cityColumn, stateColumnAndRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Extract the row number from the state column input (e.g., "P5")
  const match = stateColumnAndRow.match(/^([A-Z]+)(\d+)$/);
  if (!match) throw new Error('Invalid state column input. Use a valid range like "P5".');

  const stateColumn = match[1]; // Column letter for state
  const rowNumber = parseInt(match[2]); // Row number

  // Get city and state values from the specified columns and row
  const city = sheet.getRange(`${cityColumn}${rowNumber}`).getValue().trim();
  const state = sheet.getRange(`${stateColumn}${rowNumber}`).getValue().trim();

  if (!city || !state) return 'City or state missing.';

  // Combine city and state for geocoding
  const location = `${city}, ${state}`;

  // Use geocode.maps.co API - https://geocode.maps.co
  const apiKey = '0000000'; // Replace with your API key
  const url = `https://geocode.maps.co/search?q=${encodeURIComponent(
    location
  )}&api_key=${apiKey}`;

  try {
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());

    if (!data || data.length === 0) {
      return `No results for "${location}".`;
    }

    // Extract coordinates
    const result = data[0];
    const lat = result.lat;
    const lon = result.lon;

    // Delay to comply with API rate limits
    Utilities.sleep(1001); // Wait 1 second before allowing the next request

    return `${lat},${lon}`;
  } catch (error) {
    return `Error: ${error.message}`;
  }
}
