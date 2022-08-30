const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
var osExtensionPath = "C:/Program Files (x86)/Common Files/Adobe/CEP/extensions/com.extension.testing";

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
var csInterface = new CSInterface();

const TOKEN_PATH = csInterface.getSystemPath(SystemPath.MY_DOCUMENTS
  )+'/creds.json';
// Load client secrets from a local file.
fs.readFile(osExtensionPath+'/client_secret.json', (err, content) => {
  if (err) alert('Error loading client secret file:'+ err);
  // Authorize a client with credentials, then call the Google Sheets API.
  authorize(JSON.parse(content), mainFunc);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);
      
  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    
	
    oAuth2Client.setCredentials(JSON.parse(token));
    globalAuth = oAuth2Client;
    callback(oAuth2Client);

    
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  var csInterface = new CSInterface();
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  csInterface.openURLInDefaultBrowser(authUrl);
  var temp = prompt("Please Input the Auth code from the browser URL");
  if(temp) {
    
    oAuth2Client.getToken(temp, (err, token) => {
      if (err) return alert('Error while trying to retrieve access token'+ err+' and token = ' + token);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) alert("aaaaa" + err);
        alert('Token stored to'+ TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  }
}

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
 function mainFunc(auth) {
	const sheets = google.sheets({version: 'v4', auth});
	sheets.spreadsheets.values.get({
	  spreadsheetId: '1eQIiEhVN-bxt-pkrKUCSyTviiRCT_1s_GPO8ypPScu0',
	  range: 'Sheet1',
	}, (err, res) => {
	  if (err) return alert('The API returned an error: ' + err);
	  const rows = res.data.values;
	  alert(rows);
	});
  }


function getOS() {
 		var userAgent = window.navigator.userAgent,
 		platform = window.navigator.platform,
 		macosPlatforms = ['Macintosh', 'MacIntel', 'MacPPC', 'Mac68K'],
 		windowsPlatforms = ['Win32', 'Win64', 'Windows', 'WinCE'],
 		os = null;

 		if(macosPlatforms.indexOf(platform) != -1) {
 			os = "MAC";
 		} else if(windowsPlatforms.indexOf(platform) != -1) {
 			os = "WIN";
 		}
 		return os;
 	}