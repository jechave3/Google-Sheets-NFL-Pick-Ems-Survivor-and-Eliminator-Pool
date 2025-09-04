/** GOOGLE SHEETS FOOTBALL PICK 'EMS & SURVIVOR
 * Script Library for League Creator & Management Platform
 * v3.0
 * 09/03/2025
 * 
 * Created by Ben Powers
 * ben.powers.creative@gmail.com
 * 
 * ------------------------------------------------------------------------
 * DESCRIPTION:
 * A series of Google Apps Scripts that generate multiple sheets and a weekly Google Form
 * to be utilized for gathering pick 'ems, survivor, and eliminator selections.
 * 
 * ------------------------------------------------------------------------
 * INSTRUCTIONS:
 * make a copy and try running the Picks menu option
 * 
 * ------------------------------------------------------------
 * MENU OPTIONS WITH FUNCTION EXPLANATIONS:
 * 
 * Configuration - set up the specs for your pool
 * 
 * Member Manager - enter member names, rearrange, mark paid, revive (if using survivor/eliminator), and remove
 * 
 * Member Rename - rename a member in the sheet and update back-end name (note: this won't update the name on the form, which could cause problems, do this mid-week)
 * 
 * Form Builder - make a new form with all sorts of customization
 * 
 * Form Manager - review existing forms, turn on and off trigger to record logging, review specs of the form, copy links, etc. Also preview response count
 * 
 * Form Import - import picks for any week that has a form (only do this when you're ready to import). Should prompt to only bring in passed weeks if desired
 * ------------
 * Fetch Scores - bring in NFL outcomes to the sheet IN PROGRESS!
 * Fetch NFL Data - update the schedule data, should bring in new spreads
 * Update Formulas - should refresh formulas on all sheets that have named ranges
 * ------------
 * Enable Trigger - required for processing updates to Survivor/Eliminator evals
 * Disable Trigger - remove if causing issues or want to run without it for a while
 * ------------
 * Help & Support - opens an HTML pop-up that has a link to send me an email and this project hosted on GitHub
 * 
 * If you're feeling generous and would like to support my work,
 * here's a link to support my wife, five kiddos, and me:
 * https://www.buymeacoffee.com/benpowers
 * 
 * Thanks for checking out the script!
 * 
 * **/
 
  /**
 * Runs when the spreadsheet is opened. Checks if the script has been
 * initialized for this document and either shows the authorization card or the main menu.
 */
function onOpen() {
  // Check a specific property to see if the first-run initialization is complete.
  const docProps = PropertiesService.getDocumentProperties();
  const isInitialized = docProps.getProperty('init') === 'true';
  const config = JSON.parse(docProps.getProperty('configuration'));
  if (isInitialized) {
    let menu = SpreadsheetApp.getUi()
      .createMenu('Picks')
      .addItem('Configuration', 'launchConfiguration');
    if (docProps.getProperty('configuration')) {
      menu.addSeparator()
        .addItem('Member Manager', 'launchMemberPanel')
        .addItem('Rename a Member','showRenamePanel');
    }
    if (docProps.getProperty('members')) {
      menu.addSeparator()
        .addItem('Form Builder', 'launchFormBuilder')
        .addItem('Form Manager', 'launchFormManager')
      if (docProps.getProperty('forms')) {
        menu.addItem('Form Import', 'launchFormImport')
          .addSeparator()
          .addItem('Fetch Scores','recordWeeklyScores')
          .addItem('Update ' + LEAGUE + ' Data', 'fetchSchedule')
          .addItem('Update Formulas', 'allFormulasUpdate')
          .addItem('Deploy Sheets','setupSheets')
        if (config.pickemsInclude || config.survivorInclude) {
          menu.addSeparator()
            .addItem('Enable Triggers','createOnEditTrigger')
            .addItem('Disable Triggers','deleteOnEditTrigger');
        }
      }
    }
    if (true) {
      menu.addSeparator()
        .addItem('Check Configuration','checkDocumentConfiguration')
        .addItem('Check Members','checkDocumentMembers')
        .addItem('Delete Configuration','deleteConfiguration')
    }
      
    menu.addSeparator()
      .addItem('Help & Support','showSupportDialog')
      .addToUi();

  } else {
    SpreadsheetApp.getUi()
      .createMenu('Picks')
      .addItem('Initialize', 'launchConfiguration')
      .addToUi();
  }
}

// ============================================================================================================================================
// GLOBAL VARIABLES
// ============================================================================================================================================

// GLOBAL VARIABLES - for easy modification in the future
const LEAGUE = "NFL"; // Hopefully I"ll be able to support NCAAF at some point
const TEAMS = 32;
const REGULAR_SEASON = 18; // Regular season matchups
const WEEKS = 23; // Total season weeks (including playoffs)
const WEEKS_TO_EXCLUDE = [22]; // Break before Superbowl
const MAXGAMES = TEAMS/2;
const SCOREBOARD = 
    LEAGUE == "NFL" ? "https://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard" :
    (LEAGUE == "NCAAF" ? "https://site.api.espn.com/apis/site/v2/sports/football/college-football/scoreboard" : null);
const COLOR_PRIMARY = "#D50A0A";
const COLOR_SECONDARY = "#D50A0A";
const COLOR_TERTIARY = "#FFFFFF";

const DAY = {  0: {"name":"Sunday","index":0},  1: {"name":"Monday","index":1},  2: {"name":"Tuesday","index":2},  3: {"name":"Wednesday","index":-4},  4: {"name":"Thursday","index":-3},  5: {"name":"Friday","index":-2},  6: {"name":"Saturday","index":-1} };

const WEEKNAME = { 19: {"name":"WildCard","teams":12,"matchups":6}, 20: {"name":"Divisional","teams":8,"matchups":4}, 21: {"name":"Conference","teams":4,"matchups":2}, 23: {"name":"SuperBowl","teams":2,"matchups":1} };

const weeklySheetPrefix = "WK";
const schedulePrefix = "https://lm-api-reads.fantasy.espn.com/apis/v3/games/ffl/seasons/";
const scheduleSuffix = "?view=proTeamSchedules";
const fallbackYear = 2025;
const dayColorsObj = {"Thursday":"#fffdcc","Friday":"#e7fed1","Saturday":"#cffdda","Sunday":"#bbfbe7","Monday":"#adf7f5"};
const dayColorsFilledObj = {"Thursday":"#fffb95","Friday":"#d4ffa6","Saturday":"#abffbf","Sunday":"#89fddb","Monday":"#74f7f3"};
const dayColors = ["#fffdcc","#e7fed1","#cffdda","#bbfbe7","#adf7f5"];
const dayColorsFilled = ["#fffb95","#d4ffa6","#abffbf","#89fddb","#74f7f3"];
const configTabColor = "#ff9561";
const generalTabColor = "#aaaaaa";
const winnersTabColor = "#ffee00";
const survElimTabColors = {"survivor":"#ffee00","eliminator":"#fca503"}

const scheduleTabColor = "#472a24";

const LEAGUE_DATA = {
  "ARI": {
    "division": "NFC West",
    "division_opponents": ["LAR", "SEA", "SF"],
    "colors": [
      "#97233F",
      "#000000",
      "#FFB612",
      "#FFFFFF"
    ],
    "mascot": "🐦",
    "colors_emoji": "🔴⚫"
  },
  "ATL": {
    "division": "NFC South",
    "division_opponents": ["CAR", "NO", "TB"],
    "colors": [
      "#010101",
      "#A6192E",
      "#FFFFFF",
      "#B2B4B2"
    ],
    "mascot": "🦜",
    "colors_emoji": "⚫🔴"
  },
  "BAL": {
    "division": "AFC North",
    "division_opponents": ["CIN", "CLE", "PIT"],
    "colors": [
      "#24125F",
      "#FFFFFF",
      "#9A7611",
      "#010101"
    ],
    "mascot": "🐦‍⬛",
    "colors_emoji": "🟣🟡"
  },
  "BUF": {
    "division": "AFC East",
    "division_opponents": ["MIA", "NE", "NYJ"],
    "colors": [
      "#003087",
      "#C8102E",
      "#FFFFFF",
      "#091F2C"
    ],
    "mascot": "🐃",
    "colors_emoji": "🔵🔴"
  },
  "CAR": {
    "division": "NFC South",
    "division_opponents": ["ATL", "NO", "TB"],
    "colors": [
      "#101820",
      "#0085CA",
      "#B2B4B2",
      "#FFFFFF"
    ],
    "mascot": "🐈‍⬛",
    "colors_emoji": "⚫🔵"
  },
  "CHI": {
    "division": "NFC North",
    "division_opponents": ["DET", "GB", "MIN"],
    "colors": [
      "#091F2C",
      "#DC4405",
      "#FFFFFF"
    ],
    "mascot": "🐻",
    "colors_emoji": "🔵🟠"
  },
  "CIN": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CLE", "PIT"],
    "colors": [
      "#010101",
      "#DC4405",
      "#FFFFFF"
    ],
    "mascot": "🐅",
    "colors_emoji": "⚫🟠"
  },
  "CLE": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CIN", "PIT"],
    "colors": [
      "#311D00",
      "#EB3300",
      "#FFFFFF",
      "#EDC8A3"
    ],
    "mascot": "🟤",
    "colors_emoji": "🟤🟠"
  },
  "DAL": {
    "division": "NFC East",
    "division_opponents": ["NYG", "PHI", "WSH"],
    "colors": [
      "#0C2340",
      "#FFFFFF",
      "#87909A",
      "#7F9695"
    ],
    "mascot": "🤠",
    "colors_emoji": "🔵⚪"
  },
  "DEN": {
    "division": "AFC West",
    "division_opponents": ["KC", "LAC", "LV"],
    "colors": [
      "#0C2340",
      "#FC4C02",
      "#FFFFFF"
    ],
    "mascot": "🐴",
    "colors_emoji": "🔵🟠"
  },
  "DET": {
    "division": "NFC North",
    "division_opponents": ["CHI", "GB", "MIN"],
    "colors": [
      "#0069B1",
      "#FFFFFF",
      "#A2AAAD",
      "#010101"
    ],
    "mascot": "🦁",
    "colors_emoji": "⚪🔵"
  },
  "GB": {
    "division": "NFC North",
    "division_opponents": ["CHI", "DET", "MIN"],
    "colors": [
      "#183029",
      "#FFB81C",
      "#FFFFFF"
    ],
    "mascot": "🧀",
    "colors_emoji": "🟢🟡"
  },
  "HOU": {
    "division": "AFC South",
    "division_opponents": ["IND", "JAX", "TEN"],
    "colors": [
      "#1D1F2A",
      "#E4002B",
      "#FFFFFF",
      "#0072CE"
    ],
    "mascot": "🐂",
    "colors_emoji": "🔴🔵"
  },
  "IND": {
    "division": "AFC South",
    "division_opponents": ["HOU", "JAX", "TEN"],
    "colors": [
      "#003A70",
      "#FFFFFF",
      "#A2AAAD",
      "#1D252D"
    ],
    "mascot": "🐎",
    "colors_emoji": "🔵⚪"
  },
  "JAX": {
    "division": "AFC South",
    "division_opponents": ["HOU", "IND", "TEN"],
    "colors": [
      "#006271",
      "#D29F13",
      "#010101",
      "#9A7611"
    ],
    "mascot": "🌴",
    "colors_emoji": "🟡🔵"
  },
  "KC": {
    "division": "AFC West",
    "division_opponents": ["DEN", "LAC", "LV"],
    "colors": [
      "#C8102E",
      "#FFB81C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "🏹",
    "colors_emoji": "🔴🟡"
  },
  "LAC": {
    "division": "AFC West",
    "division_opponents": ["DEN", "KC", "LV"],
    "colors": [
      "#0072CE",
      "#FFB81C",
      "#FFFFFF",
      "#0C2340"
    ],
    "mascot": "⚡",
    "colors_emoji": "🔵🟡"
  },
  "LAR": {
    "division": "NFC West",
    "division_opponents": ["ARI", "SEA", "SF"],
    "colors": [
      "#1E22AA",
      "#FFD100",
      "#D7D2CB",
      "#FFFFFF"
    ],
    "mascot": "🐏",
    "colors_emoji": "🔵🟡"
  },
  "LV": {
    "division": "AFC West",
    "division_opponents": ["DEN", "KC", "LAC"],
    "colors": [
      "#010101",
      "#A2AAAD",
      "#FFFFFF",
      "#87909A"
    ],
    "mascot": "🏴‍☠️",
    "colors_emoji": "⚫⚪"
  },
  "MIA": {
    "division": "AFC East",
    "division_opponents": ["BUF", "NE", "NYJ"],
    "colors": [
      "#008C95",
      "#FC4C02",
      "#FFFFFF",
      "#005776"
    ],
    "mascot": "🐬",
    "colors_emoji": "🟡🔵"
  },
  "MIN": {
    "division": "NFC North",
    "division_opponents": ["CHI", "DET", "GB"],
    "colors": [
      "#582C83",
      "#FFC72C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "⚔️",
    "colors_emoji": "🟣🟡"
  },
  "NE": {
    "division": "AFC East",
    "division_opponents": ["BUF", "MIA", "NYJ"],
    "colors": [
      "#0C2340",
      "#C8102E",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🥁", //🧦
    "colors_emoji": "🔵🔴"
  },
  "NO": {
    "division": "NFC South",
    "division_opponents": ["ATL", "CAR", "TB"],
    "colors": [
      "#010101",
      "#D3BC8D",
      "#FFFFFF",
      "#A28D5B"
    ],
    "mascot": "⚜️",
    "colors_emoji": "⚫🟡"
  },
  "NYG": {
    "division": "NFC East",
    "division_opponents": ["DAL", "PHI", "WSH"],
    "colors": [
      "#001E62",
      "#A6192E",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🏗️",
    "colors_emoji": "🔵🔴"
  },
  "NYJ": {
    "division": "AFC East",
    "division_opponents": ["BUF", "MIA", "NE"],
    "colors": [
      "#115740",
      "#FFFFFF",
      "#A2AAAD",
      "#010101"
    ],
    "mascot": "✈️",
    "colors_emoji": "🟢⚪"
  },
  "PHI": {
    "division": "NFC East",
    "division_opponents": ["DAL", "NYG", "WSH"],
    "colors": [
      "#004851",
      "#E3E5E6",
      "#545859",
      "#010101"
    ],
    "mascot": "🦅",
    "colors_emoji": "🟢⚫"
  },
  "PIT": {
    "division": "AFC North",
    "division_opponents": ["BAL", "CIN", "CLE"],
    "colors": [
      "#010101",
      "#FFB81C",
      "#FFFFFF",
      "#C8102E"
    ],
    "mascot": "🏭",
    "colors_emoji": "⚫🟡"
  },
  "SEA": {
    "division": "NFC West",
    "division_opponents": ["ARI", "LAR", "SF"],
    "colors": [
      "#0C2340",
      "#78BE21",
      "#A2AAAD",
      "#FFFFFF"
    ],
    "mascot": "🌊",
    "colors_emoji": "🔵🟢"
  },
  "SF": {
    "division": "NFC West",
    "division_opponents": ["ARI", "LAR", "SEA"],
    "colors": [
      "#A6192E",
      "#B9975B",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "⛏️",
    "colors_emoji": "🔴🟡"
  },
  "TB": {
    "division": "NFC South",
    "division_opponents": ["ATL", "CAR", "NO"],
    "colors": [
      "#010101",
      "#A6192E",
      "#3D3935",
      "#DC4405"
    ],
    "mascot": "🏴‍☠️",
    "colors_emoji": "🔴⚫"
  },
  "TEN": {
    "division": "AFC South",
    "division_opponents": ["HOU", "IND", "JAX"],
    "colors": [
      "#0C2340",
      "#418FDE",
      "#B2B4B2",
      "#C8102E"
    ],
    "mascot": "🛡️",
    "colors_emoji": "🔵🔴"
  },
  "WSH": {
    "division": "NFC East",
    "division_opponents": ["DAL", "NYG", "PHI"],
    "colors": [
      "#651C32",
      "#FFB81C",
      "#FFFFFF",
      "#010101"
    ],
    "mascot": "🎖️",
    "colors_emoji": "🟤🟡"
  }
};

/**
 * Displays a custom HTML modal dialog to guide the user through authorization.
 */
function showAuthorizationCard() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Montserrat', Arial, sans-serif; padding: 10px; text-align: center; }
          h2 { color: #013369; margin-top: 0px; }
          p { font-size: 14px; color: #333; line-height: 1.6; }
          .permissions { text-align: left; background-color: #f8f9fa; border: 1px solid #e8eaed; border-radius: 8px; padding: 12px; margin-top: 15px; }
          .permissions strong { color: #D50A0A; }
          .btn { background-color: #013369; color: white; padding: 12px 24px; border: none; border-radius: 5px; cursor: pointer; font-size: 16px; font-weight: 600; margin-top: 25px; }
          .btn:hover { background-color: #2067b3; }
        </style>
      </head>
      <body>
        <h2>Welcome to NFL Picks!</h2>
        <p>Before you can get started, the scripts need your permission to run.</p>
        
        <div class="permissions">
          <strong>Why does it need these permissions?</strong>
          <ul>
            <li><strong>View and manage Sheets:</strong> To create and update this Sheet.</li>
            <li><strong>View and manager Drive:</strong> To create and update your weekly picks Forms.</li>
            <li><strong>Connect to an external service:</strong> To fetch the latest ${LEAGUE} game schedules and scores.</li>
          </ul>
        </div>

        <p style="margin-bottom: 0;">Feel free to review the code within the "Extensions" > "Apps Script" menu beforehand, but otherwise click below to start the standard Google authorization process to set up your pool.</p>

        <button class="btn" style="margin-top: 4px; padding: 10px 16px;" onclick="authorizeScript()">Let's Go!</button>

        <script>
          function authorizeScript() {
            // This button calls the server function that will trigger the auth prompt.
            google.script.run
              .withSuccessHandler(onAuthorizationSuccess)
              .withFailureHandler(onAuthorizationFailure)
              .triggerAuthorizationFlow();
          }
          
          function onAuthorizationSuccess() {
            // The success message is now a clear call to action.
            alert('Authorization successful! Please click "Configuration" from the menu again to open the config panel.');
            google.script.host.close();
          }

          function onAuthorizationFailure(error) {
            alert('Authorization failed or was canceled. Please try again. Error: ' + error.message);
          }
        </script>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(500).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'One-Time Setup Required');
}

/**
 * A simple server-side function whose only purpose is to trigger the auth flow.
 * Because this function requires a permission (accessing properties), calling it
 * will force the Google authorization dialog to appear if the user isn't yet authorized.
 */
function triggerAuthorizationFlow() {
  // This line requires permission, which is what we need.
  PropertiesService.getDocumentProperties().setProperty('init', 'true');
  
  // By setting this property, the next time onOpen runs, it will create the real menu.
  Logger.log('Script has been successfully initialized and authorized for this document.');
  onOpen();
}

/**
 * Will check for authorization, if authorized will launch configuration panel
 */
function launchConfiguration() {
  const isInitialized = PropertiesService.getDocumentProperties().getProperty('init');
  
  if (isInitialized === 'true') {
    // If we're initialized, show the real sidebar.
    configurationPanel();
  } else {
    // If this is the first run, show the authorization walk-through.
    showAuthorizationCard();
  }
}

/**
 * Creates an HTML-based configuration sidebar that mimics CardService styling.
 * This is the recommended approach for Google Sheets add-ons.
 */
function configurationPanel() {
  const html = HtmlService.createHtmlOutputFromFile('configurationPanel.html')
    .setTitle(`${LEAGUE} Pool Configuration`) 
    .setWidth(350); 
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Processes the configuration from the sidebar.
 * Uses the saveProperties() helper function.
 */
function processConfigurationSubmission(formObject) {
  const docProps = PropertiesService.getDocumentProperties();
  const previousConfig = docProps.getProperty('configuration');
  try {
    const allOptions = [
      'groupName',

      'pickemsInclude', // Pick 'Ems options below
      'pickemsAts',
      'mnfExclude',
      'commentsExclude',
      'bonusInclude',
      'mnfDouble',
      'tiebreakerInclude',
      'overUnderInclude',
      
      'survivorInclude', // Survivor options below
      'survivorStartWeek',
      'survivorAts',
      'survivorRevives',
      'survivorLives',
      
      'eliminatorInclude', // Elminator options below
      'eliminatorStartWeek',
      'eliminatorAts',
      'eliminatorRevives',
      'eliminatorLives',
      
      'customizeMatchups', // Matchup customization
      'customizeMode',
      'matchupCustomization',

      'membershipLocked', // General options below
      // 'playoffsExclude',
      'hideEmojis'
      
    ];
    const configToSave = {};
    let removeNewUserEntry = false;
    allOptions.forEach(option => {
      if (option === 'membershipLocked') {
        if (previousConfig.hasOwnProperty('membershipLocked')) {
          if (formObject.hasOwnProperty('membershipLocked')) {
            if (previousConfig.membershipLocked === false && formObject.membershipLocked === true) {
              removeNewUserEntry = true;
            }
          }
        }
      }
      if (formObject.hasOwnProperty(option)) {
        configToSave[option] = formObject[option];
      };
    });

    if (!configToSave.groupName) {
      if (configToSave.pickemsInclude) {
        configToSave.groupName = `${LEAGUE} Pick 'Ems`;
      } else {
        configToSave.groupName = `${LEAGUE} Survivor Pool`;
      }
    }
  
    let modes = ['survivor','eliminator'];
    for (const type in modes) {
      if (configToSave[`${modes[type]}Include`]) {
        if (previousConfig[`${modes[type]}StartWeek`] > configToSave[`${modes[type]}StartWeek`]) {
          Logger.log(`Previous configuration for ${type} started in week ${previousConfig[modes[type]+'StartWeek']} and new setting has moved that to ${configToSave[modes[type]+'StartWeek']}`);
          configToSave[`${modes[type]}Done`] = false;
        } else if (!configToSave[`${modes[type]}Done`]) {
          configToSave[`${modes[type]}Done`] = false;
        }
      } 
    }

    saveProperties('configuration', configToSave);
    if (removeNewUserEntry) {
      try {
        removeNewUserQuestion(fetchWeek()); // <-- Maybe should be the highest index Form created?
      }
      catch (err) {
        Logger.log(`Error removing "New User" question from form or it wasn't present: ${err.stack}`);
      }
    }
    return { success: true, data: configToSave };

  } catch (error) {
    Logger.log('Error in processConfigurationSubmission:', error);
    throw new Error('Failed to create configuration: ' + error.toString());
  }
}

/**
 * Loads the parse JSON object from Document Properties.
 * If no object is found, it returns an empty object.
 *
 * @returns {Object} The parsed object.
 */
function fetchProperties(name) {
  const string = PropertiesService.getDocumentProperties().getProperty(name);
  if (string) {
    return JSON.parse(string, (key, value) => {
      if (typeof value == 'string') {
        if (value === "true") {
          return true;
        } else if (value === "false") {
          return false;
        }
      }
      return value; // Return the value as is if not "true" or "false" string
      });
  } else {
    // If no property are found, return a default state.
    return {};
  }
}

/**
 * Saves a given JavaScript object to Document Properties as a single JSON string.
 *
 * @param {Object} object The configuration object to save.
 */
function saveProperties(name,object) {
  if (!object || typeof object !== 'object' || Object.keys(object).length === 0) {
    Logger.log('saveProperties was called with an invalid or empty object. No properties were set.');
    return; // Exit the function if there's nothing to set.
  }
  const propString = JSON.stringify(object);
  PropertiesService.getDocumentProperties().setProperty(name, propString);
}

/**
 * Deletes a single Document Property
 * If no object is found, it returns an empty object.
*/
function deleteProperties(name) {
  try {
    PropertiesService.getDocumentProperties().deleteProperty(name);
    return true;
  } catch (err) {
    Logger.log(`Error removing property ${name}: ${err.stack}`);
    return false;
  }
}


/**
 * NOT NECESSARY for final deployment
 * Remove configuration data from Document Properties
 */
function deleteConfiguration() {
  try {
    deleteProperties('configuration');
    SpreadsheetApp.getActiveSpreadsheet().toast(`SUCCESS: "configuration" removed`);
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`FAILURE: "configuration" not removed, error: ${err.stack}`);
  }
}

/**
 * Quick call to fetch via fetchProperties('configuration') with a check for existing and then providing an empty config if not found
 */
function fetchConfiguration(provideTemplate) {
  try {
    const props = fetchProperties('configuration');
    if (Object.keys(props).length > 0) {
      if (props.year) {
        if (!props.pickemsInclude) {
          props.pickemsInclude = true;
          props.tiebreakerInclude = true;
        }
        return props;
      } else {
        try {
          props.year = fetchYear(true);
        } catch (err) {
          Logger.log(`Issue fetching year when the configuration already existed within fetchConfiguration() call: ${err.stack}`);
          props.year = '2025';
        }
        return props;
      }
    } else {
      Logger.log('No existing configuration sidebar data, presenting template object');
    }
  } catch (error) {
    Logger.log('Failed to retrieve configuration sidebar data:', error);
  }
  let placeholder = {pickemsInclude:true,tiebreakerInclude:true};
  if (provideTemplate) {
    try {
      placeholder.year = fetchYear(true);
    } catch (err) {
      Logger.log(`Issue fetching year during the placeholder configuration creation within fetchConfiguration(provideTemplate) call: ${err.stack}`);
      placeholder.year = '2025';
    }
    Logger.log('Returning placeholder: ' + JSON.stringify(placeholder));
    return placeholder;
  } else {
    SpreadsheetApp.getUi().alert('No configuration found, starting configuration process...');
    showConfigDialog();
  }
}

/**
 * Retrieves all necessary data for the configuration sidebar on load.
 * Now includes a calculated message about Sunday kickoff times.
 */
function fetchConfigurationSidebarData() {
  try {
    const config = fetchConfiguration(true); // Includes Pick 'Ems ON and YEAR
    config.week = fetchWeek();
    const scriptTimeZone = Session.getScriptTimeZone();
    const formattedTime = Utilities.formatDate(new Date(), scriptTimeZone, "h:mm a',' EEE, MMM d");
    // --- [NEW LOGIC] Calculate the local kickoff time ---
    let kickoffMessage = '';
    try {
      // 1. Create a reference Date object for 1:00 PM in New York (Eastern Time).
      // We use Utilities.parseDate for a reliable way to create a date in a specific timezone.
      // The date itself doesn't matter, only the time and zone.
      const kickoffTimeET = Utilities.parseDate("13:00", "America/New_York", "HH:mm");

      // 2. Format that same moment in time for the user's detected scriptTimeZone.
      // 'h a' will format as "10 AM", "1 PM", etc.
      const localKickoffTime = Utilities.formatDate(kickoffTimeET, scriptTimeZone, "h a");

      // 3. Construct the helpful message.
      kickoffMessage = `Sunday ${LEAGUE} games will show a ${localKickoffTime} kickoff.`;

    } catch (e) {
      Logger.log("Could not calculate local kickoff time. Error: " + e.toString());
      kickoffMessage = "Could not calculate local kickoff time.";
    }
    // --- End of new logic ---
    const obj = {
      properties: config,
      league: LEAGUE,
      weekNames: WEEKNAME,
      timeZoneInfo: {
        zone: scriptTimeZone,
        currentTime: formattedTime,
        kickoffMessage: kickoffMessage
      }
    };
    return {
      properties: config,
      league: LEAGUE,
      weekNames: WEEKNAME,
      timeZoneInfo: {
        zone: scriptTimeZone,
        currentTime: formattedTime,
        kickoffMessage: kickoffMessage
      }
    };
  } catch (error) {
    Logger.log('Error preparing config sidebar data:', error);
    return {
      properties: {pickemsInclude: true},
      league: LEAGUE,
      weekNames: WEEKNAME,
      timeZoneInfo: { zone: 'Unknown', currentTime: 'N/A', kickoffMessage: '' }
    };
  }
}

function checkDocumentConfiguration() {
  try {
    const ss = fetchSpreadsheet();
    const ui = fetchUi();
    let str = '';
    const props = fetchProperties('configuration');
    Object.keys(props).forEach(key => {
      let subStr = '\n';
      if (typeof props[key] == 'object') {
        subStr += key + ': {\n';
        Object.keys(props[key]).forEach(subKey => {
          if (typeof props[key][subKey] == 'object') {
            subStr += '-' + subKey + ': {\n';
            Object.keys(props[key][subKey]).forEach(subSubKey => {
              subStr += '--' + subSubKey + ': ' + props[key][subKey][subSubKey] + '\n';
            });
            subStr += '--}\n';
          } else {
            subStr += '-' + subKey + ': ' + props[key][subKey] + '\n';
          }
        });
        subStr += '-}\n';
      } else {
          subStr += '-' + key + ': ' + props[key] + '\n';
      }
      str += (key + ': ' + (typeof props[key] === 'boolean' ? (props[key] ? '✅\n' : '❌\n') : (typeof props[key] == 'object' ? subStr : props[key] + '\n' )));
    });

    ui.alert(str,ui.ButtonSet.OK);
  } catch (err) {
    Logger.log('Failed to retrieve members sidebar data:' + err.stack);
    return { properties: {} };
  }
}


//------------------------------------------------------------------------
// SUPPORT POPUP FOR HELP - Loads HTML "supportPrompt.html" file
function showSupportDialog() {
  let html = HtmlService.createHtmlOutputFromFile('supportPrompt.html')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

//------------------------------------------------------------------------
// CONTINUATION OF SETUP - After a successful submission of the HTML prompt, this script picks up for some finishing questions and then runs the setup
function setupSheets() {
  const ss = fetchSpreadsheet();
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration'));
  
  if (!config) {
    launchConfiguration();
    ss.toast('Configuration not found or not set up yet, launching now...','⚠️ CONFIGURATION NEEDED');
    return
  }
  const memberData = JSON.parse(docProps.getProperty('members'));
  if (!memberData) {
    launchMemberPanel();
    ss.toast('Members not found or not set up yet, launching now...','⚠️ MEMBERS NEEDED');
    return;
  }
  try {
    const year = fetchYear();
    let week = fetchWeek();
    
    outcomesSheet(ss);
    Logger.log(`Deployed ${LEAGUE} Outcomes sheet`);
    if (config.pickemsInclude) {
      // Creates Weekly Totals Record Sheet
      totSheet(ss,memberData);
      Logger.log('Deployed Weekly Totals sheet');
      ss.toast('Deployed Weekly Totals sheet');

      // Creates Weekly Rank Record Sheet
      rnkSheet(ss,memberData);
      Logger.log('Deployed Weekly Rank sheet');
      ss.toast('Deployed Weekly Rank sheet');
      
      // Creates Weekly Percentages Record Sheet
      pctSheet(ss,memberData);
      Logger.log('Deployed Weekly Percentages sheet');
      ss.toast('Deployed Weekly Percentages sheet');
    
      // Creates Winners Sheet
      winnersSheet(ss,year);
      Logger.log('Deployed Winners sheet');
      ss.toast('Deployed Winners sheet');
      
      // Creates MNF Sheet
      if (!config.mnfExclude) {
        mnfSheet(ss,memberData);
        Logger.log('Deployed MNF sheet');
        ss.toast('Deployed MNF sheet');
      }
    }
    if (config.survivorInclude) {
      // Creates Survivor Sheet
      let survivor = survElimSheet(ss,config,memberData,'survivor');
      
      Logger.log('Deployed Survivor sheet');
      ss.toast('Deployed Survivor sheet');

      if (!config.pickemsInclude) {
        survivor.activate();
      }
    } else {
      try{ss.deleteSheet(ss.getSheetByName('SURVIVOR'));} catch (err) {}
    }

    if (config.eliminatorInclude) {
      // Creates Eliminator Sheet
      let eliminator = survElimSheet(ss,config,memberData,'eliminator');
      
      Logger.log('Deployed Eliminator sheet');
      ss.toast('Deployed Eliminator sheet');

      if (!config.pickemsInclude) {
        eliminator.activate();
      }
    } else {
      try{ss.deleteSheet(ss.getSheetByName('SURVIVOR'));} catch (err) {}
    }     
    
    // Creates Summary Record Sheet
    summarySheet(ss,memberData,config);
    Logger.log('Deployed Summary sheet');

    if (config.pickemsInclude) {    
      // Creates Weekly Sheets for the Current Week
      let weekly = weeklySheet(ss,week,config,memberData,false);
      Logger.log(`Deployed Weekly sheet for week ${week}`);
      ss.toast(`Deployed Weekly sheet for week ${week}`);
      weekly.activate();
    }

    ss.getSheetByName(LEAGUE).hideSheet();

    let sheet = ss.getSheetByName('Sheet1');
    if ( sheet != null ) {
      ss.deleteSheet(sheet);
    }
    Logger.log(`Deleted 'Sheet 1'`);
    
    Logger.log(`You're all set, have fun!`);
  }
  catch (err) {
    Logger.log(`runFirstStack ${err.stack}`);
  }
}

// ============================================================================================================================================
// MEMBER LIST FUNCTIONS
// ============================================================================================================================================

/**
 * Processes a revive request for a specific member and game type.
 * It fetches the current member data, updates the relevant properties for the
 * specified member, and saves the data back.
 *
 * @param {Object} data An object from the client, e.g., { memberId: "id_123...", gameType: "survivor" }.
 * @returns {Object} A success or error message object to be sent back to the client.
 */
function processReviveMember(data) {
  const { memberId, gameType, week } = data;
  
  // --- 1. Validation ---
  if (!memberId || !gameType) {
    throw new Error("Invalid request. Missing member ID or game type.");
  }
  if (gameType !== 'survivor' && gameType !== 'eliminator') {
    throw new Error(`Invalid game type: "${gameType}".`);
  }

  try {
    // --- 2. Fetch current data (Unchanged) ---
    const docProps = PropertiesService.getDocumentProperties();
    const memberData = JSON.parse(docProps.getProperty('members')) || {};
    const member = memberData.members[memberId];
    const config = JSON.parse(docProps.getProperty('configuration')) || {};

    // --- 3. Modify the data object (Now uses the 'lives array') ---
    const livesKey = gameType === 'survivor' ? 'sL' : 'eL';
    const revivesKey = gameType === 'survivor' ? 'sR' : 'eR';
    const eliminatedKey = `${gameType.substring(0,1)}O`;
    const startingLives = config[`${gameType}Lives`] || 1;
    
    // a) Reset the lives for the current week to the configured amount.
    if (!member[livesKey]) member[livesKey] = [];
    member[livesKey][week - 1] = parseInt(startingLives, 10);
    
    // b) Increment the revives used counter.
    member[revivesKey] = (member[revivesKey] || 0) + 1;
    
    // c) Clear the elimination week, as they are now back in.
    delete member[eliminatedKey];
    
     Logger.log(`Revived ${member.name} for ${gameType} in Week ${week}.`);

    // --- 4. Save the modified object back to properties (Unchanged) ---
    saveProperties('members', memberData);

    // --- 5. [THE FIX] Return the ENTIRE updated memberData object ---
    return { 
      success: true, 
      message: `${member.name} has been successfully revived!`,
      updatedMemberData: memberData // Send the fresh data back
    };

  } catch (error) {
    Logger.log(`Error in processReviveMember: `, error);
    throw new Error(`Failed to process revive: ${error.message}`);
  }
}

/**
 * Retrieves all data needed for the Member Management panel.
 * This now includes the member list AND a boolean indicating if the
 * 'Revive' feature should be displayed, based on the main pool configuration.
 *
 * @returns {Object} An object containing memberData and a showReviveButtons flag.
 */
function getMembersSidebarData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();

    // 1. Fetch the member data as before.
    const members = JSON.parse(docProps.getProperty('members')) || { order: [], details: {} }; // Ensure a default object

    // 2. Fetch the main configuration
    const config = JSON.parse(docProps.getProperty('configuration')) || {}; // Ensure a default object

    // 3. Calculate the new boolean flag based on the required conditions.
    // This will be false if either setting is false or doesn't exist.
    const showReviveSurvivorButtons = (config.survivorInclude === true && config.survivorRevives === true);
    const showReviveEliminatorButtons = (config.eliminatorInclude === true && config.eliminatorRevives === true);

    // Find a way to fold in createLivesString() to create visual of lives

    // 4. Return a single, bundled object with all the data the client needs.
    return {
      week: fetchWeek() || 1,
      memberData: members,
      showReviveSurvivorButtons: showReviveSurvivorButtons,
      showReviveEliminatorButtons: showReviveEliminatorButtons
    };

  } catch (error) {
    Logger.log('Error preparing member panel data:', error);
    // Return a safe, default structure in case of any error.
    return {
      week: fetchWeek() || 1,
      memberData: { membersOrder: [], members: {} },
      showReviveSurvivorButtons: false,
      showReviveEliminatorButtons: false
    };
  }
}

/**
 * Creates and displays the HTML modal dialog for member management.
 */
function launchMemberPanel() {
  // Create an HTML output object from a separate HTML file.
  // This is cleaner than embedding a huge string in your .gs file.
  const html = HtmlService.createHtmlOutputFromFile('memberPanel')
      .setWidth(550) // Set a comfortable width for the dialog
      .setHeight(200); // And a reasonable height
  
  // Display it as a modal dialog. The user must interact with it before returning to the sheet.
  SpreadsheetApp.getUi().showModalDialog(html, 'Member Management');
}

/**
 * Creates and displays the HTML modal dialog for member renaming
 */
function showRenamePanel() {
  const html = HtmlService.createHtmlOutputFromFile('renamePanel')
      .setWidth(400)
      .setHeight(280);
  SpreadsheetApp.getUi().showModalDialog(html, 'Rename a Member');
}

/**
 * Processes the submitted member list, compares it to the previously saved
 * state, and performs the necessary add, rename, and delete operations on the spreadsheet.
 *
 * @param {Object} newMemberData The new state of the member list from the client.
 */
function processMemberSubmission(clientData) {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const config = JSON.parse(docProps.getProperty('configuration'));
    let addedNames = [];
    // 1. Fetch the "before" state.
    const serverData = fetchProperties('members') || { memberOrder: [], members: {} };

    // 2. Process the "after" state from the client, finalizing IDs.
    const finalData = { memberOrder: [], members: {} };
    let currentWeek;
    clientData.memberOrder.forEach(id => {
      const memberDetails = clientData.members[id];
      
      if (id.startsWith('new_')) {
        const permanentId = generateUniqueId();
        finalData.memberOrder.push(permanentId);
        currentWeek = currentWeek || currentWeek() || 1;
        // --- [THE FIX] Call the new helper to create the member object ---
        finalData.members[permanentId] = createNewMember(
          memberDetails.name,
          memberDetails.paid,
          currentWeek
        );;
      } else {
        // This is an existing member. Keep their permanent ID.
        finalData.memberOrder.push(id);
        
        // Merge the old data with any new changes (like the 'paid' status).
        const existingData = serverData.members[id] || {};
        finalData.members[id] = {
          ...existingData, // Keep all old data (lives, revives, etc.)
          name: memberDetails.name, // In case of renames in the future
          paid: memberDetails.paid  // Update the paid status
        };
      }
    });
    
    // 3. Perform Deletion Logic (This part is simplified)
    // Find any IDs that were in the original serverData but are NOT in the new finalData.
    const initialIds = serverData.memberOrder || [];
    const finalIds = finalData.memberOrder;
    const deletedIds = initialIds.filter(id => !finalIds.includes(id));

    if (deletedIds.length > 0) {
        deletedIds.forEach(id => {
            const memberName = serverData.members[id]?.name || id;
            removeMemberFromSheet(memberName);
        });
    }

    // 4. Identify Additions (for future use): Names in the new list not in the old one.
    if (addedNames.length > 0) {
      Logger.log("Adding new members:", addedNames);
      // In the future, you could call an `addMemberToSheet(name)` function here.
    }
    
    //memberAddForm(addedNames);
    
    // 5. Save the new, final, authoritative state.
    saveProperties('members', finalData);
    
    return { success: true, message: 'Members updated successfully!' };

  } catch (error) {
    Logger.log("Error processing member submission:", error);
    throw new Error("Failed to update members. " + error.toString());
  }
}

/**
 * Creates a complete, correctly structured object for a new member.
 * It correctly pads the historical arrays for past weeks.
 *
 * @param {string} name The new member's name.
 * @param {boolean} isPaid The initial paid status.
 * @param {Object} config The main configuration object.
 * @param {number} joinWeek The week number the member is joining in.
 * @returns {Object} The complete new member object.
 */
function createNewMember(name, isPaid, week) {
  // --- [THE NEW LOGIC] Create padded arrays for history ---
  // Create an array with (week - 1) empty slots. (week = joining week)
  const fill = Array(week > 1 ? week - 1 : 0).fill(null);

  const newMember = {
    name: name,
    paid: isPaid,
    active: true,
    joinDate: new Date().toISOString(),
    
    // Survivor Properties
    sR: 0,
    sP: [...fill],
    sE: [...fill],
    sL: [...fill],
    sO: null, // Week Out
    // Eliminator Properties
    eR: 0,
    eP: [...fill],
    eE: [...fill],
    eL: [...fill],
    eO: null // Week Out
  };
  
  return newMember;
}


/**
 * Processes the rename submission from the renamePanel.
 * @param {Object} data An object with 'oldName' and 'newName' properties.
 */
function processRenameSubmission(data) {
  const oldName = data.oldName;
  const newName = data.newName.trim(); // Sanitize here as well for safety

  // --- Server-side validation ---
  if (!oldName || !newName || oldName === newName) {
    throw new Error("Invalid input. Please select a member and provide a different new name.");
  }
  
  const members = fetchProperties('members');
  if (!members.order || !members.details) {
    throw new Error("Could not find any member data to update.");
  }
  if (!members.order.includes(oldName)) {
    throw new Error(`The member "${oldName}" could not be found. They may have already been renamed or deleted.`);
  }
  if (members.order.map(n => n.toLowerCase()).includes(newName.toLowerCase())) {
      throw new Error(`The name "${newName}" already exists in the member list.`);
  }

  // --- Perform the update ---
  // 1. Update the 'order' array
  members.order = members.order.map(name => (name === oldName ? newName : name));
  
  // 2. Update the 'details' object
  if (members.details[oldName]) {
    members.details[newName] = members.details[oldName];
    delete members.details[oldName];
  }

  // 3. Save the updated object back to properties
  saveProperties('members',members);

  // 4. Run the function to update the spreadsheet
  renameMemberInSheet(oldName, newName);

  return { success: true, message: `Successfully renamed "${oldName}" to "${newName}".` };
}


/**
 * Finds all exact, case-sensitive matches of a member's name across all sheets
 * and removes the entire row where the name is found.
 *
 * @param {string} memberName The name of the member to remove.
 */
function removeMemberFromSheet(memberName) {
  if (!memberName) return; // Safety check

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  Logger.log(`Searching for rows to delete for member: "${memberName}"`);

  sheets.forEach(sheet => {
    const textFinder = sheet.createTextFinder(memberName)
      .matchEntireCell(true) // Crucial: Ensures "Ben" doesn't match "Benjamin"
      .matchCase(true);      // Ensures "Ben" doesn't match "ben"

    const foundCells = textFinder.findAll();
    
    if (foundCells.length > 0) {
      // We must delete from the bottom up to avoid shifting row indexes.
      foundCells.reverse().forEach(cell => {
        const row = cell.getRow();
        Logger.log(`Deleting row ${row} from sheet "${sheet.getName()}" because it contained "${memberName}"`);
        sheet.deleteRow(row);
      });
    }
  });
}

/**
 * Finds all exact, case-sensitive matches of an old member name and replaces
 * it with the new name.
 *
 * @param {string} oldName The original name to find.
 * @param {string} newName The new name to replace it with.
 */
function renameMemberInSheet(oldName, newName) {
  if (!oldName || !newName || oldName === newName) return; // Safety check

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log(`Searching for cells to rename "${oldName}" to "${newName}"`);

  // TextFinder can operate on the entire spreadsheet at once.
  const textFinder = ss.createTextFinder(oldName)
    .matchEntireCell(true) // Crucial: Fulfills your requirement
    .matchCase(true);

  // replaceAllWith is a single, efficient operation.
  const cellsReplaced = textFinder.replaceAllWith(newName);
  Logger.log(`Replaced ${cellsReplaced} instances of "${oldName}".`);
}

/**
 * NEEDS WORK - Looks for form for the week, checks what members are listed within the "eligible" entrants and adds new ones
 * 
 * @param {array} names of new players to add
 * @param {integer} week number to review
 */
function memberAddForm(names,week){
  const config = fetchConfiguration();
  const ss = fetchSpreadsheet();

  if (week == null) {
    week = fetchWeek();
  }
  if (typeof names == 'string') {
    names = [names];
  } else if (names == null) {
    names = ['New User'];
  }
  let nameQuestion, gotoPage, newUserPage, found = false;
  try {
    let formId = ss.getRangeByName('FORM_WEEK_'+week).getValue(); // FLAG
    if (formId) {
      let form = FormApp.openById(formId);
      const items = form.getItems();
      for (let a = 0; a < items.length; a++) {
        if (items[a].getType() == 'LIST' && items[a].getTitle() == 'Name') {
          nameQuestion = items[a];
          found = true;
        } else if (items[a].getType() == 'PAGE_BREAK'){
          let pageBreakItem = items[a].asPageBreakItem();
          let pageTitle = pageBreakItem.getTitle();
          if (pageTitle == 'Survivor Start') {
            gotoPage = pageBreakItem;
          } else if (pageTitle == 'New User') {
            newUserPage = pageBreakItem;
          }
        }
      }
      if (found && nameQuestion) {
        let newChoice, choices = nameQuestion.asListItem().getChoices();
        if (config.survivorInclude && survivorStart == week) { // FLAG
          try {
            for (let a = 0; a < names.length; a++) {
              if (names[a] == 'New User') {
                newChoice = nameQuestion.asListItem().createChoice(names[a],newUserPage);
                Logger.log(`New user "${names[a]}" is redirected to the "${newUserPage.getTitle()}" Form page`);
              } else {
                newChoice = nameQuestion.asListItem().createChoice(names[a],gotoPage);
                Logger.log(`New user "${names[a]}" is redirected to the "${gotoPage.getTitle()}" Form page`);
              }
              choices.unshift(newChoice);
              
            }
            nameQuestion.asListItem().setChoices(choices);
          }
          catch (err) {
            ss.toast('Issue locating survivor start question, you may need to add member manually');
            Logger.log(`memberAdd error: ${err.stack}`);
          }
        } else {
          try {
            for (let a = 0; a < names.length; a++) {
              if (names[a] == 'New User') {
                newChoice = nameQuestion.asListItem().createChoice(names[a],newUserPage);
                choices.unshift(newChoice);
                Logger.log(`New user "${names[a]}" is redirected to the "${newUserPage.getTitle()}" Form page`);
              } else {
                newChoice = nameQuestion.asListItem().createChoice(names[a],FormApp.PageNavigationType.SUBMIT);
                choices.unshift(newChoice);
                Logger.log(`New user "${names[a]}" is redirected to the submit Form page`);
              }
            }
            nameQuestion.asListItem().setChoices(choices);
          }
          catch (err) {
            ss.toast('Issue locating submit form value, you may need to add member manually');
            Logger.log(`memberAdd error: ${err.stack}`);
          }
        }
      }
    } else {
      Logger.log(`No form created yet for week ${week}, skipping addition of ${names} to form.`);
    }
  }
  catch (err) {
    Logger.log(err.stack);
    ss.toast(`Unable to add ${names} to the form.`);
  }
}


// ============================================================================================================================================
// API PULLING
// ============================================================================================================================================

// SEASON INFORMATION FUNCTIONS
//------------------------------------------------------------------------
/** 
 * Year fetching from configuration if available, if none available, fetches from ESPN API
 * 
 * @returns {int} the four digit year value of the season
 */
function fetchYear(apiPull) {
  const docProps = PropertiesService.getDocumentProperties();
  let config = JSON.parse(docProps.getProperty('configuration')) || {};
  const yearRegEx = new RegExp(/[0-9]{4}/,'g');
  let yearInvalid = true;
  if (config && !apiPull) {
    year = config.year;
    if (year) {
      return (parseInt(year).toFixed(0));;
    }
  }
  let success = false;
  if (apiPull) {
    Logger.log(`API Pull requested to fetch year`);
  } else {
    Logger.log('No year currently recorded for league, fetching from ESPN API...')
  }
  try {
    year = JSON.parse(UrlFetchApp.fetch(SCOREBOARD).getContentText()).season.year.toString();
    if (year) {
      yearInvalid = !yearRegEx.test(year);
      if (yearInvalid) {
        Logger.log(`API WARNING: Year value of "${year}" pulled from API but was not a valid year, moving on to manual submission`);
      } else {
        Logger.log(`API SUCCESS: Pulled year value of "${year}" from season information from ESPN API`)
        success = true;
      }        
    }
  } catch (err) {
    Logger.log(`API FAILURE: Unable to pull API data, moving to prompt for user... (${err})`);
  }
  if (!success) { 
    const ui = fetchUi();
    let retry = true;
    let yearPrompt = ui.prompt(`Year Entry`, `Please submit the year in a YYYY format to set the league's year for this season:`, ui.ButtonSet.OK_CANCEL);
    while (retry && yearInvalid) {
      Logger.log(JSON.stringify(yearPrompt));
      Logger.log(yearPrompt.getResponseText())
      if (yearPrompt.getSelectedButton() == ui.Button.OK) {
        year = yearPrompt.getResponseText();
        Logger.log(`Received year entry of "${year}"`);
        yearInvalid = !yearRegEx.test(year);
        if (yearInvalid) {
          yearPrompt = ui.prompt(`Retry Year Entry`, `That wasn't a valid submission, please submit the year in a YYYY format to set the league's year for this season:`, ui.ButtonSet.OK_CANCEL);
        }
      } else {
        showToast(`User canceled manual entry of year`);
        retry = false;
      }
    }
  }
  if (!yearInvalid) {
    Logger.log('Storing year value in "configuration" property...')
    config.year = year;
    try {
      saveProperties('configuration',config);
      Logger.log(`SUCCESS: Stored year key of "${year}" within the document properties.`)
      return (parseInt(year).toFixed(0));
    } catch (err) {
      Logger.log(`ERROR: Issue storing year key: ${err.stack}`);
      return null;
    }
  } else {
    return null;
  }
}

// FETCH CURRENT WEEK
function fetchWeek(negative,current) {
  let weeks, week, advance = 0;
  try {
    const obj = JSON.parse(UrlFetchApp.fetch(SCOREBOARD));
    let season = obj.season.type;
    obj.leagues[0].calendar.forEach(entry => {
      if (entry.value == season) {
        weeks = entry.entries.length;
      }
    });
    obj.events.forEach(event => {
      if (event.status.type.state != 'pre' && !current) {
        advance = 1; // At least one game has started and therefore the script will prompt for the next week
      }
    });
    let name;
    switch (season) {
      case 1:
        name = 'Preseason';
        week = obj.week.number - (weeks + 1);
        break;
      case 2: 
        name = 'Regular season';
        week = obj.week.number + advance;
        break;
      case 3:
        name = 'Postseason';
        week = obj.week.number + obj.leagues[0].calendar[1].entries.length + advance;
        break;
    }
    Logger.log(name + ' is currently active with ' + weeks + ' weeks in total, current week is: ' + week); 
    if (negative) {
      
      return week;
    } else {
      week = week <= 0 ? 1 : week;
      return week;
    }
  }
  catch (err) {
    Logger.log('ESPN API has an issue right now' + err.stack);
    return null;
  }
}

// ESPN FUNCTIONS
//------------------------------------------------------------------------
// ESPN TEAMS - Fetches the ESPN-available API data on NFL teams
function fetchTeamsESPN(year) {
  if (year == undefined) {
    year = fetchYear();
  }

  let obj = {};
  try {
    let string = schedulePrefix + year + scheduleSuffix;
    obj = JSON.parse(UrlFetchApp.fetch(string).getContentText());
    let objTeams = obj.settings.proTeams;
    return objTeams;
  }
  catch (err) {
    Logger.log('ESPN API has an issue right now');
  }  
}

// NFL TEAM INFO - script to fetch all NFL data for teams - auto for setting up trigger allows for boolean entry in column near the end
function fetchSchedule(ss,year,currentWeek,auto,overwrite) {
  // Calls the linked spreadsheet
  const timeFetched = new Date();
  ss = fetchSpreadsheet(ss);
  let all = false;
  if (currentWeek == undefined || currentWeek == null) {
    currentWeek = fetchWeek(null,true);
    all = true;
    ss.toast('Fetching complete schedule data');
  } else {
    ss.toast(`Fetching only data for week ${currentWeek}, if available.`);
  }
  // Declaration of script variables
  if (year == undefined || year == null) {
    year = fetchYear();
  }
  const objTeams = fetchTeamsESPN(year);
  const teamsLen = objTeams.length;
  let headers = ['week','date','day','hour','minute','dayName','awayTeam','homeTeam','awayTeamLocation','awayTeamName','homeTeamLocation','homeTeamName','type','divisional','division','overUnder','spread','spreadAutoFetched','timeFetched'];
  let sheetName = LEAGUE;
  let sheet, range, abbr, name, arr = [], nfl = [],espnId = [], espnAbbr = [], espnName = [], espnLocation = [], location = [], ids = [], abbrs = []; 
  
  for (let a = 0 ; a < teamsLen ; a++ ) {
    arr = [];
    if(objTeams[a].id != 0 ) {
      abbr = objTeams[a].abbrev.toUpperCase();
      name = objTeams[a].name;
      location = objTeams[a].location;
      espnId.push(objTeams[a].id);
      espnAbbr.push(abbr);
      espnName.push(name);
      espnLocation.push(location);
    }
  }

  for (let a = 0 ; a < espnId.length ; a++ ) {
    ids.push(espnId[a].toFixed(0));
    abbrs.push(espnAbbr[a]);
  }

  // Declaration of variables
  let schedule = [], home = [], dates = [], allDates = [], hours = [], allHours = [], minutes = [], allMinutes = [], byeIndex, id, date, hour, minute, weeks = Object.keys(objTeams[0].proGamesByScoringPeriod).length;
  if ( objTeams[0].byeWeek > 0 ) {
    weeks++;
  }

  location = [];
  
  for (let a = 0 ; a < teamsLen ; a++ ) {
    arr = [];
    home = [];
    dates = [];
    hours = [];
    minutes = [];
    byeIndex = objTeams[a].byeWeek.toFixed(0);
    if ( byeIndex != 0 ) {
      id = objTeams[a].id.toFixed(0);
      arr.push(abbrs[ids.indexOf(id)]);
      home.push(abbrs[ids.indexOf(id)]);
      dates.push(abbrs[ids.indexOf(id)]);
      hours.push(abbrs[ids.indexOf(id)]);
      minutes.push(abbrs[ids.indexOf(id)]);
      for (let b = 1 ; b <= weeks ; b++ ) {
        if ( b == byeIndex ) {
          arr.push('BYE');
          home.push('BYE');
          dates.push('BYE');
          hours.push('BYE');
          minutes.push('BYE');
        } else {
          if ( objTeams[a].proGamesByScoringPeriod[b][0].homeProTeamId.toFixed(0) === id ) {
            arr.push(abbrs[ids.indexOf(objTeams[a].proGamesByScoringPeriod[b][0].awayProTeamId.toFixed(0))]);
            home.push(1);
            date = new Date(objTeams[a].proGamesByScoringPeriod[b][0].date);
            dates.push(date);
            hour = date.getHours();
            hours.push(hour);
            minute = date.getMinutes();
            minutes.push(minute);
          } else {
            arr.push(abbrs[ids.indexOf(objTeams[a].proGamesByScoringPeriod[b][0].homeProTeamId.toFixed(0))]);
            home.push(0);
            date = new Date(objTeams[a].proGamesByScoringPeriod[b][0].date);
            dates.push(date);
            hour = date.getHours();
            hours.push(hour);
            minute = date.getMinutes();
            minutes.push(minute);
          }
        }
      }
      schedule.push(arr);
      location.push(home);
      allDates.push(dates);
      allHours.push(hours);
      allMinutes.push(minutes);
    }
  }
  
  // This section creates a nice table to be used for lookups and queries about NFL season
  let week, awayTeam, awayTeamName, awayTeamLocation, homeTeam, homeTeamName, homeTeamLocation, day, dayName, divisional, division, formsData = [];
  
  // Create an array of matchups per week where index of 0 is equivalent to week 1 and so forth
  let matchupsPerWeek = Array(WEEKS).fill(0);
  arr = [];
  let weekArr = [];
  for (let b = 0; b < (teamsLen - 1); b++ ) {
    for ( let c = 1; c <= weeks; c++ ) {
      if (location[b][c] == 1) {
        week = c;
        awayTeam = schedule[b][c];
        awayTeamName = espnName[espnAbbr.indexOf(awayTeam)];
        awayTeamLocation = espnLocation[espnAbbr.indexOf(awayTeam)];
        homeTeam = schedule[b][0];
        homeTeamName = espnName[espnAbbr.indexOf(homeTeam)];
        homeTeamLocation = espnLocation[espnAbbr.indexOf(homeTeam)];
        date = allDates[b][c];
        hour = allHours[b][c];
        minute = allMinutes[b][c];
        day = date.getDay();
        // Uses globalVariables.gs variable to determine day name and assign offset index
        dayName = DAY[day].name;
        day = DAY[day].index;
        divisional = LEAGUE_DATA[homeTeam].division_opponents.indexOf(awayTeam) > -1 ? 1 : 0;
        division = divisional == 1 ? LEAGUE_DATA[homeTeam].division : '';

        arr = [
          week,
          date,
          day,
          hour,
          minute,
          dayName,
          awayTeam,
          homeTeam,
          awayTeamLocation,
          awayTeamName,
          homeTeamLocation,
          homeTeamName,
          WEEKNAME.hasOwnProperty(c) ? WEEKNAME[c].name : 'Regular Season', // type
          divisional,
          division,
          '', // Placeholder for overUnder
          '', // Placeholder for spread
          '', // Placeholder for spreadAutoFetched
          timeFetched
        ];
        matchupsPerWeek[week-1] = matchupsPerWeek[week-1] + 1;
        formsData.push(arr);
      }
    }
  }

  formsData = formsData.sort((a,b) => a[1] - b[1]);
    
  for (let a = 0; a < formsData.length; a++) {
    weekArr.push(formsData[a][0]);
  }
  // Add the playoff schedule to that array of matchups per week
  Object.keys(WEEKNAME).forEach(weekNum => {
    matchupsPerWeek[weekNum-1] = WEEKNAME[weekNum].matchups;
    for (let a = 0; a < matchupsPerWeek[weekNum-1]; a++) {
      weekArr.push(parseInt(weekNum));
    }
  });

  // Create indexing array of when weeks begin and end
  let rowIndex = 2;
  let startingRow = Array(WEEKS).fill(0);
  for (let a = 1; a < startingRow.length; a++) {
    let start = 0;
    for (let b = 0; b < a; b++) {
      start = start + matchupsPerWeek[b];
    }
    startingRow[a] = 2 + start;
  }


  // Sheet formatting & Range Setting =========================
  sheet = ss.getActiveSheet();
  if ( sheet.getSheetName() == 'Sheet1' && ss.getSheetByName(sheetName) == null) {
    sheet.setName(sheetName);
  }
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    ss.insertSheet(sheetName,0);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.setTabColor(scheduleTabColor);

  sheet.setColumnWidths(1,headers.length,30);
  sheet.setColumnWidth(headers.indexOf('date')+1,60);
  sheet.setColumnWidth(headers.indexOf('dayName')+1,60);
  sheet.setColumnWidths(headers.indexOf('awayTeamLocation')+1,4,80); // All Locations & Team Names
  sheet.setColumnWidth(headers.indexOf('type')+1,110);
  sheet.setColumnWidth(headers.indexOf('division')+1,60);
  sheet.setColumnWidth(headers.indexOf('spread')+1,60);
  sheet.setColumnWidth(headers.indexOf('timeFetched')+1,110);
  range = sheet.getRange(1,1,1,headers.length);
  range.setValues([headers]);
  ss.setNamedRange(sheetName+'_HEADERS',range);

  range = sheet.getRange(1,1,weekArr.length+1,headers.length);
  range.setFontSize(8);
  range.setVerticalAlignment('middle');  
 
  ss.setNamedRange(sheetName,range);
  let rangeData = sheet.getRange(2,1,weekArr.length,headers.length);

  rangeData.setHorizontalAlignment('left');
  sheet.getRange(1,3).setNote('-4: Wednesday, -3: Thursday, -2: Friday, -1: Saturday, 0: Sunday, 1: Monday, 2: Tuesday');
  
  // Fetches sorted data
  // Sets named ranges for weekly home and away teams to compare for survivor status
  awayTeam = headers.indexOf('awayTeam')+1;
  homeTeam = headers.indexOf('homeTeam')+1;
  ss.setNamedRange(`${LEAGUE}_MATCHUPS_HEADERS`,sheet.getRange(1,1,1,headers.length));
  for (let a = 0; a < WEEKS; a++) {
    if (matchupsPerWeek[a] > 0) {
      try {
        let start = weekArr.indexOf(a+1)+2;
        let len = matchupsPerWeek[a];
        ss.setNamedRange(`${LEAGUE}_AWAY_${a+1}`,sheet.getRange(start,awayTeam,len,1));
        ss.setNamedRange(`${LEAGUE}_HOME_${a+1}`,sheet.getRange(start,homeTeam,len,1));
        ss.setNamedRange(`${LEAGUE}_MATCHUPS_${a+1}`,sheet.getRange(start,1,len,headers.length));
      }
      catch (err) {
        Logger.log(`No data entered or available for week ${a} in the spreadsheet`);
        Logger.log(err.stack);
      }
    } else {
      Logger.log(`No matchups in week ${a}`);
    }
  }
  // Sheet formatting =========================


  // Set of loops to create blank entries for playoff schedule
  const blankRow = new Array(headers.length).fill('');
  for (let a = (REGULAR_SEASON+1); a <= WEEKS; a++) {
    if (WEEKNAME.hasOwnProperty(a)) {
      for (let b = 0; b < WEEKNAME[a].matchups; b++) {
        let newRow = [...blankRow];
        newRow[0] = a; // Replace first value with week number
        formsData.push(newRow);
      }
    }
  }

  // Get scoreboard data
  const obj = JSON.parse(UrlFetchApp.fetch(SCOREBOARD));
  let scoreboardData = [];
  for (let event = 0; event < obj.events.length; event++) {
    date = new Date(obj.events[event].date);
    hour = date.getHours();
    minute = date.getMinutes();
    day = date.getDay();
    const away = obj.events[event].competitions[0].competitors.filter(x => x.homeAway === 'away')[0].team;
    const home = obj.events[event].competitions[0].competitors.filter(x => x.homeAway === 'home')[0].team;
    divisional = LEAGUE_DATA[home.abbreviation].division_opponents.indexOf(away.abbreviation) > -1 ? 1 : 0;
    division = divisional == 1 ? LEAGUE_DATA[home.abbreviation].division : '';
    let arr = [
      currentWeek,
      date,
      DAY[day].index,
      hour,
      minute,
      DAY[day].name,
      away.abbreviation,
      home.abbreviation,
      away.location,
      away.name,
      home.location,
      home.name,
      WEEKNAME.hasOwnProperty(currentWeek) ? WEEKNAME[currentWeek].name : 'Regular Season', // type
      divisional,
      division,
      (obj.events[event].competitions[0]).hasOwnProperty('odds') ? obj.events[event].competitions[0].odds[0].overUnder : '',
      (obj.events[event].competitions[0]).hasOwnProperty('odds') ? obj.events[event].competitions[0].odds[0].details : '',
      auto ? 1 : 0,
      timeFetched
    ];
    scoreboardData.push(arr);
  }
  for (let a = 0; a < formsData.length; a++) {
    if (formsData[a][0] == currentWeek) {
      formsData.splice(a,1,scoreboardData[0]);
      scoreboardData.shift();
    }
  }
  formsData.splice(formsData.indexOf(currentWeek),scoreboardData.length,...scoreboardData);

  let rows = formsData.length + 1;
  let columns = formsData[0].length;
  
  // utilities.gs functions to remove/add rows that are blank
  adjustRows(sheet,rows);
  adjustColumns(sheet,columns);
  
  let existingData = rangeData.getValues();
  const regexOverUnder = new RegExp(/^[0-9\.]+$/);
  const regexSpread = new RegExp(/^[A-Z]{2,3}\ \-[0-9\.]+$/);
  let existing = {};
  for (let a = 0; a < existingData.length; a++) {
    // Log data for each week (over/under and spread) as well as the schedule data for postseason weeks to recall later if needed
    if ((regexOverUnder.test(existingData[a][headers.indexOf('overUnder')]) || regexSpread.test(existingData[a][headers.indexOf('spread')])) || existingData[a][0] > REGULAR_SEASON) {
      let matchup = `${existingData[a][headers.indexOf('awayTeam')]}@${[existingData[a][headers.indexOf('homeTeam')]]}`;
      let rowData = existingData[a];
      existing[existingData[a][0]] = existing[existingData[a][0]] || {};
      existing[existingData[a][0]][matchup] = {};
      existing[existingData[a][0]][matchup].row = rowData;
      existing[existingData[a][0]][matchup].placed = false;
      if (existingData[a][headers.indexOf('overUnder')]) {
        existing[existingData[a][0]][matchup].overUnder = existingData[a][headers.indexOf('overUnder')];
      }
      if (existingData[a][headers.indexOf('spread')]) {
        existing[existingData[a][0]][matchup].spread = existingData[a][headers.indexOf('spread')];
      }
    }
  }

  // Checking for postseason empty slots within recently pulled data
  let missingMatchups = {};
  if (currentWeek > REGULAR_SEASON) {
    for (let a = 0; a < formsData.length; a++) {
      let formsDataWeek = formsData[a][0];
      if (formsDataWeek > REGULAR_SEASON) {
        if (formsData[a][headers.indexOf('awayTeam')] == '' || formsData[a][headers.indexOf('homeTeam') == '']) {
          missingMatchups[formsDataWeek] = missingMatchups[formsDataWeek] || {};
          missingMatchups[formsDataWeek].rows = missingMatchups[formsDataWeek].rows || [];
          missingMatchups[formsDataWeek].rows.push(a);
          missingMatchups[formsDataWeek].count = missingMatchups[formsDataWeek].count + 1 || 1;
        }
      }
    }
  }

  Object.keys(missingMatchups).forEach(week => {
    if (missingMatchups[week].count == matchupsPerWeek[week-1]) {
      Object.keys(existing[week]).forEach(matchup => {
        if (!existing[week][matchup].placed) {
          formsData[missingMatchups[week].rows[0]] = existing[week][matchup].row;
          existing[week][matchup].placed = true;
          missingMatchups[week].rows.splice(0,1);
        } else {
          Logger.log(`Already placed week ${week} matchup of ${matchup}.`);
        }
      });
    } else {
      let emptyRows = [];
      let knownMatchups = [];
      for (let a = 0; a < formsData.length; a++) {
        if (formsData[a][0] == week) {
          if (formsData[a][headers.indexOf('awayTeam')] != '' && formsData[a][headers.indexOf('homeTeam')] != '') {
            knownMatchups.push(formsData[a]);
          } else {
            emptyRows.push(a);
          }
        }
      }
      for (let a = 0; a < knownMatchups.length; a++) {
        if (existing[knownMatchups[a][0]].hasOwnProperty(`${knownMatchups[a][headers.indexOf('awayTeam')]}@${knownMatchups[a][headers.indexOf('homeTeam')]}`)) {
          existing[knownMatchups[a][0]][`${knownMatchups[a][headers.indexOf('awayTeam')]}@${knownMatchups[a][headers.indexOf('homeTeam')]}`].placed = true;
        }
      }
      Object.keys(existing[week]).forEach(matchup => {
        if (!existing[week][matchup].placed) {
          formsData.splice(emptyRows[0],1,existing[week][matchup].row);
          emptyRows.shift();
          existing[week][matchup].placed = true;
        }
      });
    }
  });
  for (let a = 0; a < formsData.length; a++ ) {
    let formsDataWeek = formsData[a][0];
    if (existing.hasOwnProperty(formsDataWeek)) {     
      if (existing[formsDataWeek].hasOwnProperty('row')) {
        Logger.log(`Replacing ${formsData[a]} with object data: ${existing[formsDataWeek].row}`);
        formsData.splice(a,1,existing[formsDataWeek].row);
      }
    }
  }

  if (Object.keys(existing).length > 0) {
    let awayIndex = headers.indexOf('awayTeam');
    let homeIndex = headers.indexOf('homeTeam');
    let spreadIndex = headers.indexOf('spread');
    let overUnderIndex = headers.indexOf('overUnder');
    for (let a = 0; a < formsData.length; a++) {
      let dataWeek = formsData[a][0];
      let matchup = `${formsData[a][awayIndex]}@${[formsData[a][homeIndex]]}`;
      if (dataWeek != currentWeek) {
        if (existing.hasOwnProperty(dataWeek)) {
          if (existing[dataWeek].hasOwnProperty(matchup)) {
            if (existing[dataWeek][matchup].hasOwnProperty('overUnder')) {
              formsData[a][overUnderIndex] = existing[dataWeek][matchup].overUnder;
            }
            if (existing[dataWeek][matchup].hasOwnProperty('spread')) {
              formsData[a][spreadIndex] = existing[dataWeek][matchup].spread;
            }
          }
        }
      }
    }
    if (!overwrite && existing.hasOwnProperty(currentWeek)) {
      let ui = fetchUi();
      let replaceAlert = ui.alert(`Found previous over/under and spread data for week ${currentWeek} in the existing NFL data. Would you like to overwright with new values?`, ui.ButtonSet.YES_NO_CANCEL);
      if (replaceAlert !== ui.Button.YES) {
        for (let a = 0; a < formsData.length; a++) {
          let dataWeek = formsData[a][0];
          let matchup = `${formsData[a][awayIndex]}@${[formsData[a][homeIndex]]}`;
          if (dataWeek === currentWeek) {
            if (existing.hasOwnProperty(dataWeek)) {
              if (existing[dataWeek].hasOwnProperty(matchup)) {
                formsData[a][overUnderIndex] = existing[dataWeek][matchup].overUnder;
                formsData[a][spreadIndex] = existing[dataWeek][matchup].spread;
              }
            }
          }
        }
      }
    }
  }
  rangeData.setValues(formsData);

  sheet.protect().setDescription(sheetName);
  try {
    sheet.hideSheet();
  }
  catch (err){
    // Logger.log('fetchSchedule hiding: Couldn\'t hide sheet as no other sheets exist');
  }
  ss.toast(`Imported all ${LEAGUE} schedule data`);
}


/**
 * A user-facing wrapper function that provides feedback and then
 * calls the main fetchSchedule logic to get the latest spreads for one week.
 *
 * @param {number} week The week number to fetch data for.
 * @returns {Object} A simple success message.
 */
function fetchLatestSpreadsForWeek(week) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    ss.toast(`Attempting to fetch latest spreads for Week ${week}...`, 'In Progress', 10);
    
    // Call your existing, powerful function with the correct parameters for a single week.
    // We pass the spreadsheet, null for year (let it auto-detect), the specific week, and false for auto, and true for automatically overriding existing spreads
    fetchSchedule(ss, null, week, false, true); 
    
    ss.toast(`Successfully updated data for Week ${week}.`, '✅ Success', 5);
    return { success: true, message: 'Data fetch complete.' };

  } catch (error) {
    Logger.log("Error in fetchLatestSpreadsForWeek: ", error);
    ss.toast(`Failed to fetch data: ${error.message}`, '❌ Error', 10);
    throw new Error(`Failed to fetch latest data. ${error.message}`);
  }
}


// NFL GAMES - output by week input and in array format: [date,day,hour,minute,dayName,awayTeam,homeTeam,awayTeamLocation,awayTeamName,homeTeamLocation,homeTeamName]
function fetchGames(week) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (week == null) {
    week = fetchWeek();
  }
  try {
    const nfl = ss.getRangeByName(LEAGUE).getValues();
    let games = [];
    for (let a = 0; a < nfl.length; a++) {
      if (nfl[a][0] == week) {
        games.push(nfl[a].slice(1));
      }
    }
    return games;
  }
  catch (err) {
    let text = 'Attempted to fetch NFL matches for week ' + week + ' but no NFL data exists, fetching now...';
    Logger.log(text);
    ss.toast(text);
    fetchSchedule(ss,null,week);
    return fetchGames(week);
  }
}

// NFL Schedule from ESPN API Scoreboard
function fetchMatchups() {
  let data = [];
  const obj = JSON.parse(UrlFetchApp.fetch(SCOREBOARD));
  let week = obj.season === 2 ? obj.week.number : (obj.season.type === 3 ? obj.week.number + REGULAR_SEASON : null);
  if (week === null) {
    throw new Error('Issue with the ESPN API for week');
  }
  for (let event = 0; event < obj.events.length; event++) {
    const date = new Date(obj.events[event].date);
    const hour = date.getHours();
    const minute = date.getMinutes();
    const day = date.getDay();
    const away = obj.events[event].competitions[0].competitors.filter(x => x.homeAway === 'away')[0].team;
    const home = obj.events[event].competitions[0].competitors.filter(x => x.homeAway === 'home')[0].team;
    const divisional = LEAGUE_DATA[home.abbreviation].division_opponents.indexOf(away.abbreviation) > -1 ? 1 : 0;
    const divName = divisional == 1 ? LEAGUE_DATA[home.abbreviation].division : '';
    const overUnder = (obj.events[event].competitions[0]).hasOwnProperty('odds') ? obj.events[event].competitions[0].odds[0].overUnder : '';
    const favorite = (obj.events[event].competitions[0]).hasOwnProperty('odds') ? obj.events[event].competitions[0].odds[0].details : '';

        // let bets = data[a][key][0].odds[0];
        // // Logger.log(data[a].shortName + ' - ' + JSON.stringify(bets));
        // arr.push([data[a].week.number,a+1,data[a].shortName.replace(" ","").replace(" ",""),
        // bets.overUnder,
        // bets.awayTeamOdds.team.abbreviation,(bets.awayTeamOdds.favorite ? parseFloat(-Math.abs(bets.spread)) : parseFloat(Math.abs(bets.spread))),
        // parseFloat(((parseFloat(bets.overUnder) + (bets.awayTeamOdds.favorite ? parseFloat(Math.abs(bets.spread)) : parseFloat(-Math.abs(bets.spread))))/2).toFixed(2)),           
        // bets.homeTeamOdds.team.abbreviation,(bets.homeTeamOdds.favorite ? parseFloat(-Math.abs(bets.spread)) : parseFloat(Math.abs(bets.spread))),
        // parseFloat(((parseFloat(bets.overUnder) + (bets.homeTeamOdds.favorite ? parseFloat(Math.abs(bets.spread)) : parseFloat(-Math.abs(bets.spread))))/2).toFixed(2))])


    let arr = [
      week,
      date,
      DAY[day].index,
      hour,
      minute,
      DAY[day].name,
      away.abbreviation,
      home.abbreviation,
      away.location,
      away.name,
      home.location,
      home.name,
      WEEKNAME.hasOwnProperty(week) ? WEEKNAME[week].name : 'Regular Season',
      divisional,
      divName,
      overUnder,
      favorite
    ];
    data.push(arr);
  }
  return data;
}

// NFL ACTIVE WEEK SCORES - script to check and pull down any completed matches and record them to the weekly sheet
function recordWeeklyScores(){
  
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration')) || {};
  const formsData = JSON.parse(docProps.getProperty('forms')) || {};
  const outcomes = fetchWeeklyScores();
  if (outcomes[0] > 0) {
    const week = outcomes[0];
    const games = outcomes[1];
    const completed = outcomes[2];
    const remaining = outcomes[3];
    const data = outcomes[4];

    const done = (games == completed);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    const pickemsInclude = config.pickemsInclude;
    const survivorInclude = config.survivorInclude;
    const eliminatorInclude = config.eliminatorInclude;
    const tiebreakerInclude = config.tiebreakerInclude;
    let outcomesRecorded = [];
    let range;
    let alert = 'CANCEL';
    if (done) {
      let text = 'WEEK ' + week + ' COMPLETE\r\n\r\nMark all game outcomes';
      if (pickemsInclude) {
        text = text + ' and tiebreaker?';
      } else {
        text = text + '?';
      }
      alert = ui.alert(text, ui.ButtonSet.OK_CANCEL);
    } else if (remaining == 1) {
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE\r\n\r\nRecord completed game outcomes?\r\n\r\n(There is one undecided game)\r\n\r\n', ui.ButtonSet.OK_CANCEL);
    } else if (remaining > 0 && remaining != games){
      alert = ui.alert('WEEK ' + week + ' INCOMPLETE\r\n\r\nRecord completed game outcomes?\r\n\r\n(There are ' + remaining + ' undecided games remaining)\r\n\r\n', ui.ButtonSet.OK_CANCEL);
    } else if (remaining == games) {
      ui.alert('WEEK ' + week + ' NOT YET STARTED\r\n\r\nNo game outcomes to record.\r\n\r\n', ui.ButtonSet.OK);
    }
    if (alert == 'OK') {
      if (pickemsInclude) {
        let sheet,matchupRange,matchups,cols,outcomeRange,outcomesRecorded,marginRange,marginRecorded,writeRange;
        try {
          sheet = ss.getSheetByName(weeklySheetPrefix+week);
          matchupRange = ss.getRangeByName(LEAGUE + '_'+week);
          matchups = matchupRange.getValues().flat();
          outcomeRange = ss.getRangeByName(LEAGUE + '_PICKEM_OUTCOMES_'+week);
          outcomesRecorded = outcomeRange.getValues().flat();
          marginRange = ss.getRangeByName(LEAGUE + '_PICKEM_OUTCOMES_'+week+'MARGIN');
          outcomesRecorded = outcomeRange.getValues().flat();
          if (tiebreakerInclude) {
            cols = matchups.length+1; // Adds one more column for tiebreaker value
          } else {
            cols = matchups.length;
          }
          writeRange = sheet.getRange(outcomeRange.getRow(),outcomeRange.getColumn(),1,cols);
          writeMarginRange = sheet.getRange(marginRange.getRow(),marginRange.getColumn(),1,cols);
        }
        catch (err) {
          const text = '❗ Issue with fetching weekly sheet or named ranges on weekly sheet, recreating now.';
          Logger.log(text + ' | ERROR: ' + err.stack);
          ss.toast(text,'MISSING WK' + week + ' SHEET');
          weeklySheet(ss,week,config,members,false);
        }
        let regex = new RegExp('[A-Z]{2,3}','g');
        let arr = [];
        for (let a = 0; a < matchups.length; a++){
          let game = matchups[a].match(regex);
          let away = game[0];
          let home = game[1];
          let outcome;        
          try {
            outcome = [];
            for (let b = 0; b < data.length; b++) {
              if (data[b][0] == away  && data[b][1] == home) {
                outcome = data[b];
              }
            }
            if (outcome.length <= 0) {
              throw new Error ('No game data for game at index ' + (a+1) + ' with teams given as ' + away + ' and ' + home);
            }
            //outcome = data.filter(game => game[0] == away && game[1] == home)[0];
            if (outcome[2] == away || outcome[2] == home) {
              if (regex.test(outcome[2])) {
                arr.push(outcome[2]);
              } else {
                arr.push(outcomesRecorded[a]);
              }
            } else if (outcome[2] == 'TIE') {
              let writeCell = sheet.getRange(outcomeRange.getRow(),outcomeRange.getColumn()+a);
              let rules = SpreadsheetApp.newDataValidation().requireValueInList([away,home,'TIE'], true).build();
              writeCell.setDataValidation(rules);
            } else {
              arr.push(outcomesRecorded[a]);
            }
          }
          catch (err) {
            Logger.log('No game data for ' + away + '@' + home);
            arr.push(outcomesRecorded[a]);
          }
          if (tiebreakerInclude) {
            try {
              if (a == (matchups.length - 1)) {
                if (outcome.length <= 0) {
                  throw new Error('No tiebreaker yet');
                }
                arr.push(outcome[3]); // Appends tiebreaker to end of array
              }
            }
            catch (err) {
              Logger.log('No tiebreaker yet');
              let tiebreakerCell = ss.getRangeByName(LEAGUE + '_TIEBREAKER_'+week);
              let tiebreaker = sheet.getRange(tiebreakerCell.getRow()-1,tiebreakerCell.getColumn()).getValue();
              arr.push(tiebreaker);
            }
          }
        }
        writeRange.setValues([arr]);
      } else if (survivorInclude || eliminatorInclude) {
        games = formsData[week].gamePlan.games;
        if (!games) {
          const text = `⚠️ Issue fetching gamePlan information for week ${week}, aborting recording of outcomes for OUTCOMES sheet`;
          Logger.log(text)
          ss.toast(text,'NO GAMEPLAN');
        } else {
          range = ss.getRangeByName(LEAGUE + '_OUTCOMES_'+week);
          outcomesRecorded = range.getValues().flat();
          let arr = [];
          for (let a = 0; a < away.length; a++) {
            arr.push([null]);
            for (let b = 0; b < data.length; b++) {
              if (data[b][0] == away[a] && data[b][1] == home[a]) {
                if (data[b][2] != null  && (outcomesRecorded[a] == null || outcomesRecorded[a] == '')) {
                  arr[a] = [data[b][2]];  
                } else {
                  arr[a] = [outcomesRecorded[a]];
                }
              }
            }        
          }
          range.setValues(arr);
        }
      }
    }
    const finishText = `✅ Recorded ${completed} game outcomes successfully`;
    Logger.log(finishText);
    ss.toast(finishText,'SUCCESS');
  } else {
    const nothing = `⛔ No outcomes to record, exiting...`
    Logger.log(nothing);
    ss.toast(nothing,`NO DATA`)
  }
}

// NFL OUTCOMES - Records the winner and combined tiebreaker for each matchup on the NFL sheet
function fetchWeeklyScores(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let obj = {};
  try{
    obj = JSON.parse(UrlFetchApp.fetch(SCOREBOARD));
  }
  catch (err) {
    Logger.log(err.stack);
    ui.alert('ESPN API isn\'t responding currently, try again in a moment.',ui.ButtonSet.OK);
    throw new Error('ESPN API issue, try later');
  }
  const season = obj.season.type;
  obj.leagues[0].calendar.forEach(entry => {
    if (entry.value == season) {
      weeks = entry.entries.length;
    }
  });
  
  if (Object.keys(obj).length > 0) {
    let games = obj.events;
    Logger.log('weeks ' + weeks + ' and games ' + games);
    let week = obj.week.number;
    let year = obj.season.year;
    Logger.log(JSON.stringify(obj));
    // Checks if preseason, if not, pulls in score data
    if (season == 1) {
      Logger.log('Regular season not yet started.\r\n\r\n Currently preseason is still underway.');
      week = week - (weeks + 1);
    //   return [0,null,null,null,null];
    } else if (season == 3) { // If in post-season, will add total weeks to output array
      week = week + REGULAR_SEASON;
    }


    let teams = [];

    // Get value for TNF being included
    let tnfInclude = true;
    try{
      tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
    }
    catch (err) {
      Logger.log('Your version doesn\'t have the TNF feature configured, add a named range "TNF_PRESENT" "somewhere on a blank CONFIG sheet cell (hidden by default) with a value TRUE or FALSE to include');
    }

    // Get existing matchup data for comparison to scores (only for TNF exclusion)
    let data = [];
    if (!tnfInclude) {
      try {
        data = ss.getRangeByName(LEAGUE).getValues();
      }
      catch (err) {
        ss.toast('No NFL data, importing now');
        fetchSchedule(ss,year);
        data = ss.getRangeByName(LEAGUE).getValues();
      }
      for (let a = 0; a < data.length; a++) {
        if (data[a][0] == week && (tnfInclude || (!tnfInclude && data[a][2] >= 0))) {
          teams.push(data[a][6]);
          teams.push(data[a][7]);
        }
      }
    }
    // Loop through games provided and creates an array for placing
    let all = [];
    let count = 0;
    let away, awayScore,home, homeScore,tiebreaker,winner,competitors;
    for (let a = 0; a < games.length; a++){
      let outcomes = [];
      awayScore = '';
      homeScore = '';
      tiebreaker = '';
      winner = '';
      competitors = games[a].competitions[0].competitors;
      away = (competitors[1].homeAway == 'away' ? competitors[1].team.abbreviation : competitors[0].team.abbreviation);
      home = (competitors[0].homeAway == 'home' ? competitors[0].team.abbreviation : competitors[1].team.abbreviation);
      if (games[a].status.type.completed) {
        if (tnfInclude || (!tnfInclude && (teams.indexOf(away) >= 0 || teams.indexOf(home) >= 0))) {
          count++;
          awayScore = parseInt(competitors[1].homeAway == 'away' ? competitors[1].score : competitors[0].score);
          homeScore = parseInt(competitors[0].homeAway == 'home' ? competitors[0].score : competitors[1].score);
          tiebreaker = awayScore + homeScore;
          winner = (competitors[0].winner ? competitors[0].team.abbreviation : (competitors[1].winner ? competitors[1].team.abbreviation : 'TIE'));
          outcomes.push(away,home,winner,tiebreaker);
          all.push(outcomes);
        }
      }      
    }
    // Sets info variables for passing back to any calling functions
    let remaining = games.length - count;
    let completed = games.length - remaining;

    // Outputs total matches, how many completed, and how many remaining, and all matchups with outcomes decided;
    return [week,games.length,completed,remaining,all];
  } else {
    Logger.log('ESPN API returned no games');
    ui.alert('ESPN API didn\'t return any game information. Try again later and make sure you\'re checking while the season is active',ui.ButtonSet.OK);
  }
}

// LEAGUE LOGOS - Saves URLs to logos to a Document Property variable named "logos" {CURRENTLY UNUSED}
function fetchLogos(){
  let obj = {};
  let logos = {};
  try{
    obj = JSON.parse(UrlFetchApp.fetch(SCOREBOARD));
  }
  catch (err) {
    Logger.log(err.stack);
    ui.alert('ESPN API isn\'t responding currently, try again in a moment.',ui.ButtonSet.OK);
    throw new Error('ESPN API issue, try later');
  }
  
  if (Object.keys(obj).length > 0) {
    let games = obj.events;
    // Loop through games provided and creates an array for placing
    for (let a = 0; a < games.length; a++){
      let competitors = games[a].competitions[0].competitors;
      let teamOne = competitors[0].team.abbreviation;
      let teamTwo = competitors[1].team.abbreviation;
      let teamOneLogo = competitors[0].team.logo;
      let teamTwoLogo = competitors[1].team.logo;
      logos[teamOne] = teamOneLogo;
      logos[teamTwo] = teamTwoLogo;
    }
    Logger.log(logos);
    const docProps = PropertiesService.getDocumentProperties();
    try {
      let logoProp = docProps.getProperty('logos');
      let tempObj = JSON.parse(logoProp);
      if (Object.keys(tempObj).length < nflTeams) {
        docProps.setProperty('logos',JSON.stringify(logos));
      }
    }
    catch (err) {
      Logger.log('Error fetching logo object, creating one now');
      docProps.setProperty('logos',JSON.stringify(logos));
    }
  }
  return logos;
}

// ============================================================================================================================================
// MENU FUNCTIONS
// ============================================================================================================================================

// CREATE MENU - this is the standard setup once the sheet has been configured and the data is all imported
function createMenu(lock,trigger) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (lock == undefined || lock == null) {
    lock = membersSheetProtected();
  }
  let pickems = false;
  try{
    pickems = ss.getRangeByName('PICKEMS_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('Issue gathering PICKEMS_PRESENT cell, you may not have completed setup correctly.');
    pickems = true;
  }
  let tnfInclude = true;
  try{
    tnfInclude = ss.getRangeByName('TNF_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('Issue gathering TNF_PRESENT cell, you may not have completed setup correctly.');
  }
  let bonus = false;
  try{
    bonus = ss.getRangeByName('BONUS_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('Issue gathering BONUS_PRESENT cell, you may not have completed setup correctly.');
  }
  let mnfDouble = false;
  try{
    mnfDouble = ss.getRangeByName('MNF_DOUBLE').getValue();
  }
  catch (err) {
    Logger.log('Issue gathering MNF_DOUBLE cell, you may not have completed setup correctly.');
  }
  let menu = ui.createMenu('Picks');
    menu.addItem('Create a Form','formCreateAuto')
      .addItem('Open Current Form','openForm');
  if (pickems) {
    menu.addItem('Week Sheet Creation','weeklySheetCreate');
  }
  menu.addSeparator();
  if (tnfInclude) {
    menu.addItem('Check Responses','formCheckAlert')
      .addItem('Import Thursday Picks','dataTransferTNF')
      .addItem('Import Picks','dataTransfer');
  } else {
    menu.addItem('Check Responses','formCheckAlert')
      .addItem('Import Picks','dataTransfer');
  }
  menu.addSeparator()
    .addItem('Check ' + LEAGUE + ' Scores','recordWeeklyScores')
    .addItem('Update ' + LEAGUE + ' Schedule', 'fetchSchedule');
  menu.addSeparator();
  if (!bonus) {
    menu.addItem('Enable Bonus','bonusUnhide');
  } else if (mnfDouble) {
    menu.addSubMenu(ui.createMenu('Bonus')
      .addItem('Hide Game Bonus Value Row','bonusHide')
      .addItem('MNF Double Value Disable','bonusDoubleMNFDisable')
      .addItem('Random Game of the Week','bonusRandomGameSet'));
  } else {
    menu.addSubMenu(ui.createMenu('Bonus')
      .addItem('Hide Game Bonus Value Row','bonusHide')
      .addItem('MNF Double Value Enable','bonusDoubleMNFEnable')
      .addItem('Random Game of the Week','bonusRandomGameSet'));
  }
  menu.addSeparator();
  if (!lock) {
  menu.addItem('Add Member(s)','memberAdd')
    .addItem('Remove Member','memberRemove')
    .addItem('Lock Members','createMenuLocked');
  } else {
    menu.addItem('Reopen Members','createMenuUnlocked');
  }
  menu.addSeparator();
  menu.addItem('Refresh Formulas','allFormulasUpdate')
    .addItem('Help & Support','showSupportDialog')
    .addToUi();
  if (trigger) {
    deleteOnOpenTriggers();
    let id = ss.getId();
    ScriptApp.newTrigger('createMenu')
      .forSpreadsheet(id)
      .onOpen()
      .create();
  }
}


// ============================================================================================================================================
// FORM FUNCTIONS
// ============================================================================================================================================

function toggleConfig() {
  const docProps = PropertiesService.getDocumentProperties();
  if (docProps.getProperty('configuration') == null) {
    const settings = docProps.getProperty('settings');
    docProps.setProperty('configuration',settings);
    docProps.deleteProperty('settings');
  } else {
    const configuration = docProps.getProperty('configuration');
    docProps.setProperty('settings',configuration);
    docProps.deleteProperty('configuration');
  }
}


/**
 * [CORRECTED & BULLETPROOF] Gathers and aggregates all data for the dashboard.
 * This version is resilient to missing or incomplete data properties.
 */
function getFormDashboardData() {
  try {
    const allProps = PropertiesService.getDocumentProperties().getProperties();
    
    // --- [THE FIX] Provide robust defaults for all major properties ---
    const config = JSON.parse(allProps['configuration'] || '{}');
    const memberData = JSON.parse(allProps['members'] || '{ "memberOrder": [], "members": {} }');
    const formsObject = JSON.parse(allProps['forms'] || '{}');
    
    const apiWeek = fetchWeek(null, true);
    
    let forms = [];
    
    // 1. Find all form data from the forms object
    for (const week in formsObject) {
      // Use optional chaining (?.) for safe property access
      const form = formsObject[week];
      if (!form) continue; // Skip if the form entry is null for some reason

      // Create a new object to avoid modifying the original
      const formsData = { 
        week: parseInt(week),
        ...form
      };
      // 2. Augment the data, using safe fallbacks for every value
      try {
        formsData.isActive = FormApp.openById(formsData.formId).isAcceptingResponses();
      } catch (err) {
        formsData.isActive = false;
      }
      
      formsData.respondents = form.respondents;
      formsData.responseCount = form.respondents.length;
      formsData.nonRespondents = form.nonRespondents;

      formsData.membershipLocked = formsData.gamePlan?.membershipLocked || false;
      formsData.newMembers = form.newMembers;
      formsData.gameCount = formsData.gamePlan?.games?.length || 0;
      
      formsData.pickemsInclude = formsData.gamePlan?.pickemsInclude || false;
      formsData.pickemsAts = formsData.gamePlan?.pickemsAts || false;
      formsData.survivorInclude = formsData.gamePlan?.survivorInclude  || false;
      formsData.survivorAts = formsData.gamePlan?.survivorAts  || false;
      formsData.eliminatorInclude = formsData.gamePlan?.eliminatorInclude  || false;
      formsData.eliminatorAts = formsData.gamePlan?.eliminatorAts  || false;;
      forms.push(formsData);
    }
    
    forms.sort((a, b) => a.week - b.week);
    
    // --- Calculation logic with safe defaults ---
    const totalMembers = memberData.memberOrder?.length || 0;
    const totalResponses = forms.reduce((acc, f) => acc + f.responseCount, 0);
    const totalPossibleResponses = totalMembers * forms.length;
    
    return {
      groupName: config.groupName || `${LEAGUE} Picks Pool`,
      forms: forms,
      totalMembers: totalMembers,
      apiWeek: apiWeek,
      overallResponseRate: totalPossibleResponses > 0 ? (totalResponses / totalPossibleResponses) : 0
    };
  } catch (error) {
    Logger.log("Error in getFormDashboardData: ", error);
    // Ensure we throw an error that the client can parse
    throw new Error("Could not load form data. " + error.message);
  }
}

/**
 * The main function to launch the forms panel
 */
function launchFormManager() {
    const html = HtmlService.createHtmlOutputFromFile('formManager').setWidth(1000).setHeight(115);
    SpreadsheetApp.getUi().showModalDialog(html, 'Form Manager');
}

/**
 * Toggles the 'isAcceptingResponses' status of a Google Form
 * Also manages the onFormSubmit trigger for that form.
 * @param {string} formId The ID of the form to toggle.
 * @returns {Object} An object containing the new active status.
 */
function toggleFormStatus(formId) {
  try {
    const form = FormApp.openById(formId);
    const currentState = form.isAcceptingResponses();
    const newState = !currentState;
    form.setAcceptingResponses(newState);

    // --- [THE NEW LOGIC] ---
    // Find any existing triggers for THIS specific form and delete them.
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(trigger => {
      if (trigger.getTriggerSourceId() === formId) {
        ScriptApp.deleteTrigger(trigger);
        Logger.log(`Deleted existing trigger for form ID: ${formId}`);
      }
    });

    if (newState === true) {
      // If we are ACTIVATING the form, create a new onFormSubmit trigger.
      ScriptApp.newTrigger('handleFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
      Logger.log(`Created new onFormSubmit trigger for form ID: ${formId}`);
    }
    
    return { success: true, newStatus: newState };
  } catch (error) {
    Logger.log(`Failed to toggle status for form ${formId}:`, error);
    throw new Error(`Could not update form status. ${error.message}`);
  }
}

/**
 * Enables or disables the onFormSubmit trigger for a specific form.
 * @param {string} formId The ID of the form to manage the trigger for.
 * @param {boolean} shouldBeEnabled The desired state for the trigger.
 * @returns {Object} A success message.
 */
function setFormSubmitTrigger(formId, shouldBeEnabled) {
  try {
    // 1. Clean up any existing triggers for this form.
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(trigger => {
      if (trigger.getTriggerSourceId() === formId) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    let toastMessage = ''; // Variable to hold our feedback message

    if (shouldBeEnabled) {
      // 2a. If enabling, create a new trigger.
      const form = FormApp.openById(formId);
      ScriptApp.newTrigger('handleFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
      toastTitle = `✅ TRIGGER ADDED`;
      toastMessage = `Auto-sync trigger has been created for the form.`;
      Logger.log(`Auto-sync trigger ENABLED for form ID: ${formId}`);
    } else {
      // 2b. If disabling, we've already deleted the trigger.
      toastTitle = `❌ TRIGGER DELETED`;
      toastMessage = `Auto-sync trigger has been removed for the form.`;
      Logger.log(`Auto-sync trigger DISABLED for form ID: ${formId}`);
    }

    // 3. Store the preference (unchanged).
    const formsData = fetchProperties('forms');
    const week = getWeekFromFormId(formId); // Your existing helper
    if (week && formsData[week]) {
      formsData[week].autoSync = shouldBeEnabled;
      saveProperties('forms', formsData);
    }
    
    // --- [THE FIX] ---
    // 4. Display the toast message.
    SpreadsheetApp.getActiveSpreadsheet().toast(toastMessage,toastTitle);

    return { success: true, newStatus: shouldBeEnabled };
  } catch (error) {
    Logger.log(`Failed to set trigger for form ${formId}:`, error);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${error.message}`, '❌ FAILED', 10);
    throw new Error(`Could not update trigger. ${error.message}`);
  }
}

/** 
 * A reverse-lookup to find a week by form ID.
*/
function getWeekFromFormId(formId) {
  const formsData = fetchProperties('forms');
  for (const week in formsData) {
    if (formsData[week].formId === formId) {
      return week;
    }
  }
  return null;
}

/**
 * onFormSubmit trigger will call this function
 * It's a simple wrapper that determines the week and calls our main sync function.
 * @param {Object} e The event object passed by the onFormSubmit trigger.
 */
function handleFormSubmit(e) {
  try {
    // 1. Get the Form object and its unique ID from the event object.
    const form = e.source;
    const submittedFormId = form.getId();
    let week = null;

    // 2. Look up the form ID in our 'forms' property.
    const formsData = fetchProperties('forms'); // Your existing helper
    if (formsData && Object.keys(formsData).length > 0) {
      for (const weekNum in formsData) {
        if (formsData[weekNum].formId === submittedFormId) {
          week = parseInt(weekNum, 10);
          break; // We found our match, no need to loop further.
        }
      }
    }

    // 3. Fallback to parsing the title if the lookup fails (optional but safe).
    if (!week) {
      Logger.log(`Could not find form ID ${submittedFormId} in the 'forms' property. Falling back to parsing title.`);
      const formTitle = form.getTitle();
      const weekMatch = formTitle.match(/Week (\d+)/);
      if (weekMatch && weekMatch[1]) {
        week = parseInt(weekMatch[1], 10);
      }
    }

    // 4. If we have a week, run the main sync function.
    if (week) {
      Logger.log(`Form submit detected for Week ${week}. Running sync...`);
      // Run our main, robust sync function.
      syncFormResponses(week);
    } else {
      // This is a critical error, as we couldn't identify the form.
      Logger.log(`CRITICAL: Could not determine week for submitted form with ID: ${submittedFormId} and Title: "${form.getTitle()}". Sync aborted.`);
    }
  } catch (error) {
    // Log any errors that occur during the sync process itself.
    Logger.log(`An error occurred during the onFormSubmit sync: ${error.stack}`);
  }
}

/**
 * The main controller for the form creation process.
 */
function launchFormBuilder() {
  const docProps = PropertiesService.getDocumentProperties();
  if (docProps.getProperty('configuration') == null) {
    Logger.log(`No configuration data present, please begin by configuring the pool`);
    const ui = SpreadsheetApp.getUi();
    if (ui.alert(`Configuration Missing`,`No configuration data found, please run the "Configuration" function first before building your first sheet.\n\nClick OK to go there now.`,ui.ButtonSet.OK_CANCEL) === ui.Button.OK) {
      launchConfiguration();
      return;
    } else {
      Logger.log('User declined to create configuration to being processing form creation');
      showToast('Unable to start form creation due to no configuration file')
      return;
    }
  }
  let openEnrollment = false;
  const config = JSON.parse(docProps.getProperty('configuration'));
  if (docProps.getProperty('members') == null) {
    Logger.log(`No members data present.`);
    const ui = SpreadsheetApp.getUi();
    if (ui.alert(`Members Missing`,`No members data found, please run the "Member Management" function to add some participants, otherwise the form will default to providing a name prompt for all users${config.membershipLocked ? ' (you currently have membership locked and this will unlock membership.)': '.'}\n\nClick OK to go there now.`,ui.ButtonSet.OK_CANCEL) === ui.Button.OK) {
      launchMemberPanel();
      return;
    } else {
      if (config.membershipLocked) {
        Logger.log(`Unlocking membership due to no members supplied prior to first form creation.`)
        config.membershipLocked = false;
        saveProperties('configuration',config);
      }
      openEnrollment = true;
      Logger.log('User declined to create members to being processing form creation');
    }
  }
  const isFirstRun = docProps.getProperty('hasCreatedFirstForm') !== 'true';
  try {
    if (isFirstRun || !checkFileExists(docProps.getProperty('templateId'))) {
      docProps.setProperty('hasCreatedFirstForm','false');
      handleFirstFormCreation();
    } else {
      const htmlTemplate = HtmlService.createTemplateFromFile('formCreatorPanel');
      const htmlOutput = htmlTemplate.evaluate().setWidth(700).setHeight(210);
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, `Create Form${openEnrollment ? ' - Open Enrollment' : ''}`);
    }
  } catch (err) {
    Logger.log(`Error starting form creation: ${err.stack}`)  
  }
}

/**
 * Encapsulates the entire first-run and theme customization workflow.
 */
function handleFirstFormCreation() {
  const ui = SpreadsheetApp.getUi();
  const docProps = PropertiesService.getDocumentProperties();
  try {
    const templateForm = getTemplateForm();
    Logger.log('Template Form ' + templateForm) ;
    if (!templateForm) return;

    let response = ui.alert(
      'Customize Theme (One Time Only)',
      'Before creating your first weekly form, would you like to open the template to customize the colors and header image?',
      ui.ButtonSet.YES_NO
    );
    
    docProps.setProperty('hasCreatedFirstForm', 'true');
    
    if (response === ui.Button.YES) {
      showLinkDialog(templateForm.getEditUrl(), 'Template Customization', 'Open Your Template',`Once you've optionally updated the header image and color palette, restart the form creation function. Note: changing the header image should automatically modify the color palette`);
    }
  } catch (err) {
    if (err.message.includes("CANCELLED_BY_USER")) {
      showToast('Form creation cancelled.');
    } else {
      ui.alert('An unexpected error occurred: ' + err.message);
    }
  }
}

/**
 * Function to gether necessary inputs for form creation pop-up
 */
function fetchFormCreationData() {
  try {
    const docProps = PropertiesService.getDocumentProperties();
    const configuration = JSON.parse(docProps.getProperty('configuration'));
    const earliestWeek = fetchWeek() || 1;
    const scheduleAnalysis = analyzeScheduleData(earliestWeek);
    let apiWeek = null;
    try {
      apiWeek = fetchWeek(null, true)
    } catch (err) {
      Logger.log(`Error in fetching API week for fetchFormCreationData function: ${err.stack}`);
    }
    return {
      configuration: configuration,
      matchupData: {
        available: scheduleAnalysis.available,
        matchups: scheduleAnalysis.matchups
      },
      validitySummary: scheduleAnalysis.validitySummary,
      apiWeek: apiWeek,
      leagueData: LEAGUE_DATA
    };
  } catch (error) {
    Logger.log("A critical error occurred in fetchFormCreationData: ", error);
    // Explicitly return a safe, default object on failure.
    return { 
      configuration: {}, 
      matchupData: {
        available: false, matchups: [], },
      validitySummary: {},
      apiWeek: apiWeek,
      leagueData: LEAGUE_DATA
    };
  }
}
/**
 * Fetches, filters, and analyzes schedule data to create a
 * week-by-week data quality summary.
 * @param {number} earliestWeek - The first week to include in the analysis.
 * @returns {Object} An object containing the filtered matchups and the validity summary.
 */
function analyzeScheduleData(earliestWeek) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LEAGUE); // Assumes LEAGUE is a global const like 'NFL'
    
    if (!sheet) throw new Error(`Sheet '${LEAGUE}' not found.`);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { available: false, matchups: [], validitySummary: {} };
    
    const headers = data.shift();
    const weekCol = headers.indexOf('week');
    const dateCol = headers.indexOf('date');
    const timeFetchedCol = headers.indexOf('timeFetched');
    const spreadCol = headers.indexOf('spread');
    const overUnderCol = headers.indexOf('overUnder');
    const autoFetchedCol = headers.indexOf('spreadAutoFetched');

    // 1. Filter data to relevant weeks and convert to objects
    const matchups = data
      .filter(row => row[weekCol] >= earliestWeek)
      .map(row => {
        let matchupObject = {};
        headers.forEach((header, index) => {
          let value = row[index];
          if ((index === dateCol || index === timeFetchedCol) && value instanceof Date) {
            matchupObject[header] = value.toISOString();
          } else {
            matchupObject[header] = value;
          }
        });
        return matchupObject;
      });

    if (matchups.length === 0) return { available: true, matchups: [], validitySummary: {} };

    // 2. Group matchups by week for analysis
    const gamesByWeek = matchups.reduce((acc, game) => {
      const week = game.week;
      if (!acc[week]) acc[week] = [];
      acc[week].push(game);
      return acc;
    }, {});

    // 3. Analyze each week's data to create the summary
    const validitySummary = {};
    for (const week in gamesByWeek) {
      const weekGames = gamesByWeek[week];
      let firstTimeFetched = weekGames[0]?.timeFetched || null;
      
      const analysis = {
        auto: weekGames.every(g => g[headers[autoFetchedCol]] === true || g[headers[autoFetchedCol]] === 1),
        timeFetched: firstTimeFetched,
        spreads: weekGames.every(g => g[headers[spreadCol]] !== ''),
        overUnders: weekGames.every(g => g[headers[overUnderCol]] !== '')
      };

      // Check for inconsistent fetch times
      const isTimeConsistent = weekGames.every(g => g.timeFetched === firstTimeFetched);
      if (!isTimeConsistent) {
        analysis.timeFetched = 'ERROR';
      }
      
      validitySummary[week] = analysis;
    }

    return {
      available: true,
      matchups: matchups,
      validitySummary: validitySummary
    };
  } catch (e) {
    Logger.log("Error in analyzeScheduleData: ", e);
    return { available: false, matchups: [], validitySummary: {} };
  }
}

/**
 * Updates the 'Schedule' sheet with user-provided override data.
 * @param {number} week The week being updated.
 * @param {Object} customData The object of user edits from the client.
 */
function updateScheduleData(week, customData) {
  if (!customData || Object.keys(customData).length === 0) {
    return; // Nothing to update
  }
  
  Logger.log(`Applying ${Object.keys(customData).length} user overrides for Week ${week}.`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LEAGUE); // e.g., 'NFL'
  if (!sheet) throw new Error(`Sheet '${LEAGUE}' not found.`);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const weekCol = headers.indexOf('week');
  const awayCol = headers.indexOf('awayTeam');
  const homeCol = headers.indexOf('homeTeam');
  const spreadCol = headers.indexOf('spread');
  const overUnderCol = headers.indexOf('overUnder');
  const timeFetchedCol = headers.indexOf('timeFetched');

  let changesMade = 0;
  
  // Loop through the data array (skipping headers)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[weekCol] == week) {
      const gameId = `${row[awayCol]}@${row[homeCol]}`;
      if (customData[gameId]) {
        const edits = customData[gameId];
        if (edits.spread !== undefined) {
          data[i][spreadCol] = edits.spread;
          changesMade++;
        }
        if (edits.overUnder !== undefined) {
          data[i][overUnderCol] = edits.overUnder;
          changesMade++;
        }
        // Also update the timeFetched and auto-fetch status to reflect manual edit
        data[i][timeFetchedCol] = new Date();
        // Assuming you have a 'spreadAutoFetched' column
        const autoFetchedCol = headers.indexOf('spreadAutoFetched');
        if (autoFetchedCol > -1) {
          data[i][autoFetchedCol] = 0; // It's now a manual override
        }
      }
    }
  }

  // If changes were made, write the entire data range back to the sheet.
  if (changesMade > 0) {
    sheet.getDataRange().setValues(data);
    Logger.log("Successfully updated the Schedule sheet.");
  }
}

/**
 * Receives the "game plan" from the UI, performs pre-flight checks,
 * confirms with the user if necessary, and then executes the form creation.
 *
 * @param {Object} gamePlan The detailed object of user intent from the client.
 */
function createNewFormForWeek(gamePlan) {
  const ui = fetchUi();
  const week = gamePlan.week;
  const year = gamePlan.year;
  // 1. Pre-flight Checks
  const docProps = PropertiesService.getDocumentProperties();
  const forms = JSON.parse(docProps.getProperty(`forms`)) || {};
  const config = JSON.parse(docProps.getProperty('configuration'));

  
  let warnings = [];
  if (forms[week]) {
    if (forms[week].formId) {
      warnings.push("An existing form for this week will be deleted.\n\nAny associated responses in the database file will be archived.");
    }
  }

  // 2. User Confirmation (if necessary)
  if (warnings.length > 0) {
    const message = "WARNING:\n\n" + warnings.join("\n") + "\n\nAre you sure you want to proceed?";
    const response = ui.alert(message, ui.ButtonSet.YES_NO);
    if (response !== ui.Button.YES) {
      throw new Error("Form creation canceled by user.");
    }
  }

  try {
    // 3. Prepare for Creation
    if (forms[week] && forms[week].formId) {
      try {
        DriveApp.getFileById(forms[week].formId).setTrashed(true);
      } catch (e) {
        Logger.log(`Could not trash old form with ID ${forms[week].formId}. It may have been deleted already. Continuing.`);
      }
    }
    // Clear out old data from properties
    delete forms[week];

    // 4. Execute the "Worker"
    const newFormDetails = buildFormFromGamePlan(gamePlan);
    
    // Record the current state of these properties for data fetching integrity
    gamePlan.pickemsInclude = config.pickemsInclude;
    gamePlan.survivorInclude = newFormDetails.survivorInclude;
    gamePlan.eliminatorInclude = newFormDetails.eliminatorInclude;
    gamePlan.pickemsAts = config.pickemsAts;
    gamePlan.survivorAts = config.survivorAts;
    gamePlan.eliminatorAts = config.eliminatorAts;

    // 5. Save the new state
    const newFormsData = {
      formId: newFormDetails.formId,
      editUrl: newFormDetails.editUrl,
      publishedUrl: newFormDetails.publishedUrl,
      active: true,
      respondents: [],
      nonRespondents: newFormDetails.eligibleMembers,
      gamePlan: gamePlan
    };
    forms[week] = newFormsData;
    forms[week].imported = false;

    // 6. Storing properties
    saveProperties('forms',forms);
    
    // 7. Setting up synce for form
    try {
      setFormSubmitTrigger(newFormDetails.formId, true) 
    } catch (err) {
      Logger.log(`Could not set up trigger for new week ${week} form: ${err.stack}`);
    }
    
    showLinkDialog(newFormDetails.publishedUrl, `Success!`, `Open Week ${week} Form`, `Open '${gamePlan.formName}' and review if desired. Then share with your group! Keep in mind you have it set to ${gamePlan.membershipLocked ? 'allow new members to join via the form' : 'prevent new members from joining via the form'}`);

    return { success: true, message: `✅ Successfully created form for week ${week}.` };
    
  } catch (err) {
    Logger.log(`Failed to create form for week ${week}:`, err.stack);
    throw new Error(`Failed to create form: ${err.stack}`);
  }
}

/**
 * Fetches, validates, or creates the single backend "database" spreadsheet.
 * This is the reciprocal to your formTemplate function.
 *
 * @returns {Spreadsheet} The active, valid Spreadsheet object for the database.
 */
function getDatabaseSheet() {
  const docProps = PropertiesService.getDocumentProperties();
  const dbId = docProps.getProperty('databaseId');

  if (dbId) {
    try {
      // Try to open the spreadsheet using the stored ID.
      const file = DriveApp.getFileById(dbId);
      if (file.getName() && !file.isTrashed()) {
        return SpreadsheetApp.openById(dbId);
      } else {
        docProps.deleteProperty('databaseId');
      }
    } catch (e) {
      // This catch block runs if the file was deleted or permissions changed.
      Logger.log(`Could not open database sheet with ID "${dbId}". It may have been deleted. A new one will be created.`);
      docProps.deleteProperty('databaseId'); // Clear the invalid ID
    }
  }

  // If we reach here, either there was no ID or the old one was invalid.
  const ui = SpreadsheetApp.getUi();
  ui.alert('Backend Database Not Found', 'A private spreadsheet for storing form responses could not be found. A new one will be created now.', ui.ButtonSet.OK);
  
  const config = JSON.parse(docProps.getProperty('configuration'));
  const formsFolder = getFormsFolder(config.groupName || `${LEAGUE} Picks Pool`);

  const dbName = `DO NOT DELETE - ${fetchYear()} ${config.groupName || LEAGUE + ' Picks Pool - Form Responses'}`;
  const newDb = SpreadsheetApp.create(dbName);
  const newDbId = newDb.getId();
  
  // Move to the forms folder as well
  DriveApp.getFileById(newTemplateId).moveTo(formsFolder);
  
  // Add a helpful note for the user in the new sheet.
  newDb.getSheets()[0].getRange('A1').setValue(`This sheet is the private backend for the ${config.groupName} pool. Please do not delete, rename, or share this file.`);

  docProps.setProperty('databaseId', newDbId);
  Logger.log(`Created new database sheet with ID: ${newDbId}`);
  
  return newDb;
}

/**
 * Fetches, validates, or creates the template form.
 * Now prompts the user with a YES/NO choice if the template needs to be created.
 *
 * @returns {Form|null,boolean} The Form object, or null if the user cancels and a boolean if the form ID stored didn't work upon onset
 */
function getTemplateForm() {
  Logger.log('Beginning process to get template form...');
  const docProps = PropertiesService.getDocumentProperties();
  const templateId = docProps.getProperty('templateId');
  if (templateId) {
    Logger.log('Found template id: ' + templateId);
    try {
      const file = DriveApp.getFileById(templateId);
      if (file && !file.isTrashed()) {
        Logger.log('Found an existing template file, attempting to open...');
        return FormApp.openById(templateId);
      }
    } catch (err) {
      Logger.log(`Could not open template form with ID "${templateId}". ${err.stack}`);
      docProps.deleteProperty('templateId');
    }
  }
  const templateName = 'PICKS TEMPLATE - Customize Header and Color';
  const newTemplate = FormApp.create(templateName);
  const newTemplateId = newTemplate.getId();
  
  // Ensure the template is moved into the correct folder.
  const config = fetchProperties('configuration');
  const formsFolder = getFormsFolder(config.groupName || `${LEAGUE} Picks Pool`);
  DriveApp.getFileById(newTemplateId).moveTo(formsFolder);

  docProps.setProperty('templateId', newTemplateId);
  Logger.log(`Created new template form with ID: ${newTemplateId}`);

  return newTemplate;
}

/**
 * Fetches, validates, or creates the folder for storing weekly forms.
 *
 * @param {string} groupName The name of the pool, passed in to avoid extra fetches.
 * @returns {Folder} The active, valid Folder object.
 */
function getFormsFolder(groupName) {
  const docProps = PropertiesService.getDocumentProperties();
  const folderId = docProps.getProperty('folderId');
  
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      // A simple check to ensure it's a valid folder.

      if (folder.getName() && !folder.isTrashed()) {
        return folder;
      }
    } catch (err) {
      Logger.log(`Could not open forms folder with ID "${folderId}". It may have been deleted. A new one will be created.${err.stack}`);
      docProps.deleteProperty('folderId');
    }
  }

  const folderName = `${fetchYear()} ${groupName ? groupName : LEAGUE + ' Picks Pool'} Forms`;
  const newFolder = DriveApp.createFolder(folderName);
  docProps.setProperty('folderId', newFolder.getId());
  Logger.log(`Created new forms folder: "${folderName}"`);
  
  return newFolder;
}

/**
 * Takes a final, validated "game plan" and builds a Google Form.
 * This function is now modular and uses the ID-based members object.
 *
 * @param {Object} gamePlan The detailed plan for the form.
 * @returns {Object} An object with the new form's ID and URLs.
 */
function buildFormFromGamePlan(gamePlan) {
  try {
    // --- Setup and Variable Initialization ---
    const docProps = PropertiesService.getDocumentProperties();
    const config = JSON.parse(docProps.getProperty('configuration'));
    const memberData = JSON.parse(docProps.getProperty('members'));
    const week = gamePlan.week;
    const formName = gamePlan.formName;

      // 1. Get the validated helper objects.
    const formsFolder = getFormsFolder(config.groupName || `${LEAGUE} Picks Pool`);
    const templateForm = getTemplateForm();
    let databaseSheet = getDatabaseSheet();
    const databaseSheetId = databaseSheet.getId();
    
    // 2. Create the new form by copying the template.
    const newFormFile = DriveApp.getFileById(templateForm.getId()).makeCopy(formName, formsFolder);
    const form = FormApp.openById(newFormFile.getId());

    const urlFormEdit = form.shortenFormUrl(form.getEditUrl());
    const urlFormPub = form.shortenFormUrl(form.getPublishedUrl());

    try {
      // Get the initial count of sheets before linking
      const initialSheets = databaseSheet.getSheets();
      const initialSheetNames = initialSheets.map(sheet => sheet.getName());

      // Set the form's destination. This creates a new sheet in the spreadsheet.
      form.setDestination(FormApp.DestinationType.SPREADSHEET, databaseSheet.getId());
      let waitPeriod = 1;
      let updatedSheets = [];
      let updatedSheetNames = [];
      let foundNewSheet = false;
      while (waitPeriod <= 40 && !foundNewSheet) {
        // Wait a moment for Google to create the new sheet...
        Utilities.sleep(500);
        SpreadsheetApp.flush();
        
        // Re-fetch the spreadsheet to get the current state
        databaseSheet = SpreadsheetApp.openById(databaseSheetId);
        
        // Get the updated list of sheets
        updatedSheets = databaseSheet.getSheets();
        updatedSheetNames = updatedSheets.map(sheet => sheet.getName());
        
        Logger.log(`Attempt ${waitPeriod}: Found ${updatedSheets.length} sheets: [${updatedSheetNames.join(', ')}]`);
        
        // Check if we have a new sheet
        if (updatedSheets.length > initialSheets.length) {
          Logger.log(`Success! Found new sheet after ${waitPeriod} attempts.`);
          foundNewSheet = true;
          break; // Exit the loop immediately
        }
        
        waitPeriod++;
      }
      if (!foundNewSheet) {
        throw new Error(`Timed out waiting for new response sheet to be created after ${waitPeriod - 1} attempts.`);
      }

    
      let newResponseSheet = null;


      // Method 1: Find the sheet that wasn't there before
      for (const sheet of updatedSheets) {
        if (!initialSheetNames.includes(sheet.getName())) {
          newResponseSheet = sheet;
          break; // Exit as soon as we find it
        }
      }
        
      // Method 2: Fallback - look for sheets with "Form Responses" pattern
      if (!newResponseSheet) {
        Logger.log("Fallback: Looking for 'Form Responses' pattern...");
        for (const sheet of updatedSheets) {
          if (sheet.getName().includes('Form Responses')) {
            // Check if this one is new by comparing against initial list
            if (!initialSheetNames.includes(sheet.getName())) {
              newResponseSheet = sheet;
              break;
            }
          }
        }
      }
      if (!newResponseSheet) {
        // Debug: Show exactly what we found
        Logger.log("DEBUG - Initial sheets: " + JSON.stringify(initialSheetNames));
        Logger.log("DEBUG - Updated sheets: " + JSON.stringify(updatedSheetNames));
        Logger.log("DEBUG - Difference: " + JSON.stringify(updatedSheetNames.filter(name => !initialSheetNames.includes(name))));
        throw new Error("Could not identify the newly created response sheet.");
      }
      Logger.log(`Found new response sheet: "${newResponseSheet.getName()}"`);

      // Now let's rename and organize it
      const newSheetName = `WK${week}`;
      
      // Check if a sheet with the desired name already exists
      const existingSheet = databaseSheet.getSheetByName(newSheetName);
      
      if (existingSheet) {
        Logger.log(`An existing sheet named '${newSheetName}' was found. Archiving it now.`);
        
        let archiveIndex = 1;
        let archiveName = `WK${week}_ARCHIVE${archiveIndex}`;
        
        // Loop to find a unique archive name that doesn't already exist
        while (databaseSheet.getSheetByName(archiveName) !== null) {
          archiveIndex++;
          archiveName = `WK${week}_ARCHIVE${archiveIndex}`;
        }
        
        // Rename and hide the old sheet
        existingSheet.setName(archiveName);
        existingSheet.hideSheet();
        
        Logger.log(`Successfully archived old sheet as '${archiveName}'.`);
      }
      
      // Now safely rename the newly linked sheet
      newResponseSheet.setName(newSheetName);
      newResponseSheet.activate(); // Brings the new tab to the front
      
      Logger.log(`Successfully linked form and renamed response sheet to '${newSheetName}'.`);
    } catch (err) {
      // If linking fails, delete the form to avoid orphans
      DriveApp.getFileById(form.getId()).setTrashed(true);
      Logger.log(`Failed to link form to spreadsheet and manage tab. Form deleted. Error: ${err.stack}`);
      throw new Error("Could not link form to the backend database. Please check permissions.");
    }

    const nameQuestion = form.addListItem().setTitle('Select Your Name').setRequired(true);

    // --- Build Pick'em Questions (if applicable) ---
      if (config.pickemsInclude) {
      // 1. Add a Section Header to act as the title for this part of the form.
      form.addSectionHeaderItem().setTitle("🏈 Weekly Pick 'Em Selections");
      // 2. If ATS is enabled for Pick'em, add a very clear instructional message.
      if (config.pickemsAts) {
        form.addSectionHeaderItem()
          .setTitle('🔢 Instructions: Pick Against the Spread (ATS)')
          .setHelpText('For each game, select the team you believe will win WITH the point spread. The point spread is listed in the help text of each question.');
      }
      buildPickemQuestions(form, gamePlan, config);
    } else {
      Logger.log(`No pick 'ems pool active, moving on to survivor/eliminator`)
    }

    const finalSubmitPage = form.addPageBreakItem().setGoToPage(FormApp.PageNavigationType.SUBMIT);
    
    // --- [THE NEW COMBINED LOGIC] ---
    const pageDestinations = {}; // Will map memberId -> destination page
    
    const survivorIsActive = !config.survivorDone && config.survivorInclude && week >= config.survivorStartWeek;
    const eliminatorIsActive = !config.eliminatorDone && config.eliminatorInclude && week >= config.eliminatorStartWeek;

    const sLS = config.survivorLives;
    const eLS = config.eliminatorLives;
    let allSurvivorTeamsForWeek = buildTeamList(gamePlan, config, config.survivorAts);
    let allEliminatorTeamsForWeek = buildTeamList(gamePlan, config, config.eliminatorAts)
    // 1. First, find members active in BOTH contests and create their combined page.
    if (survivorIsActive && eliminatorIsActive) {
      Logger.log('Creating possible destinations for instances where both Survivor and Eliminator are both active')
      memberData.memberOrder.forEach(memberId => {
        const member = memberData.members[memberId];
        if (member && member.active && member.sL[week-1] > 0 && member.eL[week-2] > 0) {
          let survivorHelp = sLS == 1 ? `One Survivor Life: ${createLivesString(member.sL[week-2],sLS)}` : `Survivor Lives: ${createLivesString(member.sL[week-2],sLS)} (${member.sL[week-2] < sLS ? member.sL[week-2] + ' remaining' : 'all remaining'})`;
          let eliminatorHelp = eLS == 1 ? `One Eliminator Life: ${createLivesString(member.eL[week-2],eLS)}` : `Eliminator Lives: ${createLivesString(member.eL[week-2],eLS)} (${member.eL[week-2] < eLS ? member.eL[week-2] + ' remaining' : 'all remaining'})`;
          const helpText = `${survivorHelp}  |  ${eliminatorHelp}`;
          const combinedPage = form.addPageBreakItem().setTitle(`${member.name}'s Picks`).setHelpText(helpText);
          combinedPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);
          
          // Add both questions to this single page
          addContestQuestion(form, 'survivor', member, config.survivorAts, config.survivorStartWeek, allSurvivorTeamsForWeek);
          addContestQuestion(form, 'eliminator', member, config.eliminatorAts, config.eliminatorStartWeek, allEliminatorTeamsForWeek);
          
          pageDestinations[memberId] = combinedPage; // Assign their destination
        }
      });
    }  
    // 2. Next, handle members active in ONLY ONE contest.
    Logger.log('Creating possible destinations members in only one of the Survivor or Eliminator contests')
    memberData.memberOrder.forEach(memberId => {
      // Skip members we've already handled
      if (pageDestinations[memberId]) return;

      const member = memberData.members[memberId];
      if (member && member.active) {
        if (survivorIsActive && member.sL > 0) {
          const helpText = `Survivor Lives: ${createLivesString(member.sL[week-2],sLS)} (${member.sL[week-2]})`;
          const survivorPage = form.addPageBreakItem().setTitle(`${member.name}'s Survivor Pick`).setHelpText(helpText);
          survivorPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);
          addContestQuestion(form, 'survivor', member, config.survivorAts, config.survivorStartWeek, allSurvivorTeamsForWeek);
          pageDestinations[memberId] = survivorPage;

        } else if (eliminatorIsActive && member.eL > 0) {
          const helpText = `Eliminator Lives: ${createLivesString(member.eL[week-2],eLS)} (${member.eL[week-2]})`;
          const eliminatorPage = form.addPageBreakItem().setTitle(`${member.name}'s Eliminator Pick`).setHelpText(helpText);
          eliminatorPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);
          addContestQuestion(form, 'eliminator', member, config.eliminatorAts, config.eliminatorStartWeek, allEliminatorTeamsForWeek);
          pageDestinations[memberId] = eliminatorPage;
        }
      }
    });

    // 3. Finally, build the name dropdown using our new destination map.
    Logger.log('Creating links to name drop-down based on members enrollment in Survivor and/or Eliminator pools')
    let nameChoices = [], eligibleMembers = [];
    memberData.memberOrder.forEach(memberId => {
      const member = memberData.members[memberId];
      if (member && member.active) {
        // If a destination was calculated for them, use it. Otherwise, default to the final submit page.
        const destination = pageDestinations[memberId] || finalSubmitPage;
        
        // Prevent adding members who have no questions to answer
        if (destination === finalSubmitPage && !config.pickemsInclude) return;
        
        nameChoices.push(nameQuestion.createChoice(member.name, destination));
        eligibleMembers.push(member.name);
      }
    });
    
    // Add 'New User' option if applicable
    if (!config.membershipLocked) {
      Logger.log('🔓 Membership is unlocked--creating a new user question');
      const newUserPage = buildNewUserPage(form, config, gamePlan);
      nameChoices.unshift(nameQuestion.createChoice(('✏️ NEW USER'), newUserPage));

    } else {
      Logger.log('🔒 Membership is locked--no new user question added');
    }
    
    Logger.log('📝 Setting Name Choices');
    nameQuestion.setChoices(nameChoices);

    // --- Final Touches ---
    Logger.log('↩️ Returning information to form creation controller...');
    return {
      formId: form.getId(),
      editUrl: urlFormEdit,
      publishedUrl: urlFormPub,
      eligibleMembers: eligibleMembers,
      survivorInclude: survivorIsActive,
      eliminatorInclude: eliminatorIsActive
    };
  } catch (err) {
    Logger.log(`❗ Encountered an issue during the creation of the form: ${err.stack}`)
    Logger.log(`❌ Deleting form if it was created somewhere during the process...`);
    try {
      DriveApp.getFileById(form.getId()).setTrashed(true);
      Logger.log(`🗑 Successfully trashed form.`)
    } catch (err) {
      Logger.log(`❌ Unable to trash new form or it didn't exist yet.`)
    }
    return err;
  }  
}

/**
 * Adds a single, correctly filtered contest question to the form.
 */
function addContestQuestion(form, contestType, member, isAts, startWeek, allTeamsForWeek) {
  const picksKey = contestType === 'survivor' ? 'sP' : 'eP';
  const allMemberPicks = member[picksKey] || [];
  const relevantPicks = allMemberPicks.slice(startWeek - 1);
  
  const availableTeams = allTeamsForWeek.filter(team => {
    const teamAbbr = team.split(' ')[0];
    return !relevantPicks.includes(teamAbbr);
  });
  
  const baseTitle = (contestType === 'survivor' ? '👑 ' : '💀 ') + capitalize(contestType) + (contestType === 'eliminator' ? ' Loser' + (isAts ? ' ATS' : '') + ' Pick' : ' Winner' + (isAts ? ' ATS' : '') + ' Pick');
  const uniqueTitle = `${baseTitle} (${member.name})`;
  const helpText = `Select which team you believe will ${isAts ? (contestType === 'survivor' ? 'WIN' : 'LOSE') + ' when factoring in the given spread' : (contestType === 'survivor' ? 'WIN' : 'LOSE') + ' this week'}.`;
  form.addListItem()
    .setTitle(uniqueTitle)
    .setHelpText(helpText)
    .setChoiceValues(availableTeams)
    .setRequired(true);
}

/**
 * Builds all Pick'em related questions on the form.
 */
function buildPickemQuestions(form, gamePlan, config) {
  Logger.log("Building Pick'em questions...");
  gamePlan.games.forEach(game => {
    let item = form.addMultipleChoiceItem();
    const evening = game.hour >= 17;
    const mnf = evening && game.dayName === "Monday";
    let title = `${game.awayTeamLocation} ${game.awayTeamName} at ${game.homeTeamLocation} ${game.homeTeamName}${game.divisonal == 1 ? ' ('+game.division+' Divsional Game)':''}`;
    let helpText = `${mnf ? 'Monday Night Football' : game.dayName} at ${formatTime(game.hour, game.minute)}`;
    if (game.spread) helpText += `  |  Spread: ${game.spread}`;
    if (game.bonus > 1) title += ` (${game.bonus}x Bonus)`;
    if (config.tiebreakerInclude) {
      tiebreakerMatchup = `${game.awayTeamLocation} ${game.awayTeamName} at ${game.homeTeamLocation} ${game.homeTeamName}`;
      tiebreakerOverUnder = game.overUnder;
    }
    item.setTitle(title)
      .setHelpText(helpText)
      .setChoices([
        item.createChoice(`${!config.hideEmojis ? ' ' + LEAGUE_DATA[game.awayTeam].mascot: ''} ${game.awayTeam}`), // + LEAGUE_DATA[game.awayTeam].colors_emoji 
        item.createChoice(`${!config.hideEmojis ? ' ' + LEAGUE_DATA[game.homeTeam].mascot: ''} ${game.homeTeam}`)]) // + LEAGUE_DATA[game.homeTeam].colors_emoji 
      .showOtherOption(false)
      .setRequired(true);
  });
  if (config.tiebreakerInclude) { // Excludes tiebreaker question if tiebreaker is disabled
    let numberValidation = FormApp.createTextValidation()
      .setHelpText('Input must be a whole number between 0 and 120')
      .requireWholeNumber()
      .requireNumberBetween(0,120)
      .build();
      // Tiebreaker question
    let helpText = `Total points combined between ${tiebreakerMatchup}${config.overUnderInclude && tiebreakerOverUnder > 0 ? ' (current betting line is ' + tiebreakerOverUnder + ')' : ''}`
    form.addTextItem()
      .setTitle('Tiebreaker')
      .setHelpText(helpText)
      .setRequired(true)
      .setValidation(numberValidation);
  }
  if(config.commentsInclude) { // Excludes comment question if comments are disabled
    form.addTextItem()
      .setTitle('Comments')
      .setHelpText('Passing thoughts...');
  }  
}

/**
 * Builds the page for a new user to sign up.
 * This function now intelligently adds Survivor and/or Eliminator questions
 * directly to the same page if it's the first week of the contest.
 *
 * @param {Form} form The Google Form object to add items to.
 * @param {Object} config The main configuration object for the pool.
 * @param {Object} gamePlan The game plan for the current week.
 * @returns {PageBreakItem} The created page break item for navigation.
 */
function buildNewUserPage(form, config, gamePlan) {
  // 1. Create the page break and set its final destination.
  const newUserPage = form.addPageBreakItem().setTitle('New User Signup');
  newUserPage.setGoToPage(FormApp.PageNavigationType.SUBMIT);

  // 2. Add the "Name" input question.
  const nameValidation = FormApp.createTextValidation()
    .setHelpText('Enter a minimum of 2 characters, up to 30.')
    .requireTextMatchesPattern(".{2,30}") // A simpler, more reliable regex for length
    .build();
    
  form.addTextItem()
    .setTitle('Enter Your Name')
    .setHelpText('Please enter your name as it will appear in the pool.')
    .setRequired(true)
    .setValidation(nameValidation);

  // --- [THE NEW LOGIC] ---
  // 3. Conditionally add contest questions directly to this page.  
  
  // Check if it's the first week for the Survivor pool
  const survivorIsActiveFirstWeek = !config.survivorDone && config.survivorInclude && gamePlan.week == config.survivorStartWeek;
  if (survivorIsActiveFirstWeek) {
    Logger.log(`New user section ELIGIBLE for survivor start week (${gamePlan.week}), adding queston...`);
    const sLS = config.survivorLives;
    const isAts = config.survivorAts;
    let survivorHelp = `Select the team you believe will WIN${isAts ? ' AGAINST THE SPREAD':''}.  |  `;
    survivorHelp += sLS == 1 ? `One Survivor Life: ${createLivesString(sLS)}` : `Survivor Lives: ${createLivesString(sLS)} (${member.sL[week-2] < sLS ? member.sL[week-2] + ' remaining' : 'all remaining'})`;    
    const allTeamsForWeek = buildTeamList(gamePlan, config, isAts);
    form.addListItem()
      .setTitle(`Survivor ${isAts ? "ATS" : "" } Pick`)
      .setHelpText(survivorHelp)
      .setChoiceValues(allTeamsForWeek)
      .setRequired(true);
  } else {
    Logger.log(`New user section INELIGIBLE for survivor start week (${gamePlan.week})`);
  }

  // Check if it's the first week for the Eliminator pool
  const eliminatorIsActiveFirstWeek = !config.eliminatorDone && config.eliminatorInclude && gamePlan.week == config.eliminatorStartWeek;
  if (eliminatorIsActiveFirstWeek) {
    Logger.log(`New user section ELIGIBLE for eliminator start week (${gamePlan.week}), adding queston...`);
    const eLS = config.eliminatorLives;
    const isAts = config.eliminatorAts;
    let eliminatorHelp = `Select the team you believe will LOSE${isAts ? ' AGAINST THE SPREAD':''}.  |  `;
    eliminatorHelp += eLS == 1 ? `One Eliminator Life: ${createLivesString(eLS)}` : `Eliminator Lives: ${createLivesString(eLS)} (${member.eL[week-2] < eLS ? member.eL[week-2] + ' remaining' : 'all remaining'})`;    
    const allTeamsForWeek = buildTeamList(gamePlan, config, isAts);
    form.addListItem()
      .setTitle(`Eliminator ${isAts ? "ATS" : "" } Pick`)
      .setHelpText(eliminatorHelp)
      .setChoiceValues(allTeamsForWeek)
      .setRequired(true);
  } else {
    Logger.log(`New user section INELIGIBLE for eliminator start week (${gamePlan.week})`);
  }
  
  // 4. Return the page break item so the main function can use it for navigation.
  return newUserPage;
}

/**
 * Takes all games within gameplan and populates an array with the team abbreviations, including emojis if enabled
 * 
 * @param {object} all data submitted by form creation panel
 * @param {object} league data with necessary booleans/values
 * @param {boolean} optional value to determin if spreads should be added
 * @returns {array} all teams selected for the given week to provide in a survivor/eliminator drop-down with ATS
 */
function buildTeamList(gamePlan, config, isAts) {
  isAts = isAts || (config.survivorAts || config.elminatorAts); // Fallback
  return gamePlan.games.flatMap(game => {
    const awayEmojis = `${!config.hideEmojis ? ' ' + LEAGUE_DATA[game.awayTeam].mascot : ''}`;
    const homeEmojis = `${!config.hideEmojis ? ' ' + LEAGUE_DATA[game.homeTeam].mascot : ''}`;
    if (!isAts || game.spread === 'PK' || game.spread == 0) {
      return [
        `${game.awayTeam}${awayEmojis}`,
        `${game.homeTeam}${homeEmojis}`];
    }
    // Extract the numeric value and determine which team is favored
    const spreadMatch = game.spread.match(/([A-Z]+)\s*(-?\d+\.?\d*)/);
    if (!spreadMatch) return [`${game.awayTeam}${awayEmojis}`,`${game.homeTeam}${homeEmojis}`]; // Fallback
    
    const [,favoriteTeam, spreadValue] = spreadMatch;
    const numericSpread = parseFloat(spreadValue);
    
    // Determine spreads for away and home teams
    const awaySpread = favoriteTeam === game.awayTeam ? numericSpread : Math.abs(numericSpread);
    const homeSpread = favoriteTeam === game.homeTeam ? numericSpread : Math.abs(numericSpread);
    
    return [
      `${game.awayTeam} ${awaySpread > 0 ? '+' : ''}${awaySpread}${awayEmojis}`,
      `${game.homeTeam} ${homeSpread > 0 ? '+' : ''}${homeSpread}${homeEmojis}`
    ];
  });
}

/**
 * Combines emoji or square entries to present user with visual of remaining lives
 * 
 * @param {boolean} state of whether emojis are present or not
 * @param {integer} remaining lives per user
 * @param {integer} total lives by default
 * @returns {string} character string representing lives left and those have been lost (red dot or empty square)
 */
function createLivesString(remaining, total) {
  return '🟢'.repeat(remaining)+'⚫'.repeat(total - remaining);
}

/**
 * Checks for existence of a file based on given ID
 * 
 * @param {string} file ID to check
 * @returns {boolean} true if file exists, false if not
 */
function checkFileExists(fileId) {
  // Check if fileId is valid (not null, undefined, or empty string)
  if (!fileId || typeof fileId !== 'string' || fileId.trim() === '') {
    return false;
  }  
  try {
    // DriveApp.getFileById works for both files AND folders
    const file = DriveApp.getFileById(fileId);
    if (file && !file.isTrashed()) {
      return true;
    } else {
      return false;
    }
  } catch (error) {
    return false; // Exception means file/folder doesn't exist or isn't accessible
  }
}

/**
 * Checks for existence of a folder based on given ID
 * 
 * @param {string} folder ID to check
 * @returns {boolean} true if folder exists, false if not
 */
function checkFolderExists(folderId) {
  if (!folderId || typeof folderId !== 'string' || folderId.trim() === '') {
    return false;
  }
  try {
    const folder = DriveApp.getFolderById(folderId);
    if (folder && !folder.isTrashed()) {
      return true;
    } else {
      return false;
    }
  } catch (error) {
    return false;
  }
}


function capitalize(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

// Simple function to convert 24 hour entry into AM/PM string with minutes
function formatTime(hour,minute) {
  const suffix = hour >= 12 ? 'PM' : 'AM';
  return `${hour > 12 ? hour - 12 : hour}:${minute < 10 ? '0' + minute : minute} ${suffix}`;
}

function deleteDatabaseId() {
  deleteProperties('databaseId');
}
function deleteFormTemplateId() {
  deleteProperties('templateId');
}
function deleteFormFolderId() {
  deleteProperties('folderId');
}
function deleteFormFolderOld() {
  deleteProperties('folder');
}
function deleteFormsDataProp() {
  deleteProperties('forms');
}
function deleteConfigurationProp() {
  deleteProperties('configuration');
}
function testGetFormsFolder() {
  Logger.log(getFormsFolder());
}
function getTemplateFormTest() {
  const docProps = PropertiesService.getDocumentProperties();
  const templateId = docProps.getProperty('hasCreatedFirstForm');
  Logger.log(templateId);
}
function deleteFirstFormCheck() {
  deleteProperties('hasCreatedFirstForm');
}





/**
 * [NEW] The main data-gathering function for the Import Picks panel.
 * This is called by the client-side script on load.
 */
function getFormImportData(week) {
  try {
    week = week || fetchWeek();
    // 1. Run sync first to get the latest respondent metadata.
    const syncResult = syncFormResponses(week);
    // 2. Fetch the clean, final data needed for the panel.
    const docProps = PropertiesService.getDocumentProperties();
    const formsData = JSON.parse(docProps.getProperty('forms'));
    const memberData = JSON.parse(docProps.getProperty('members'));

    const config = JSON.parse(docProps.getProperty('configuration'));
    
    const allCreatedWeeks = Object.keys(formsData).map(Number).sort((a,b) => a-b);
    week = week || allCreatedWeeks[allCreatedWeeks.length - 1] || fetchWeek(null, true);
    const gamePlan = formsData[week]?.gamePlan;

    if (!gamePlan) {
      throw new Error(`Could not find a 'gamePlan' for Week ${week}.`);
    }

    // 3. Determine which games are upcoming.
    const matchups = getInvalidPickMatchups();
    
    // 4. Determine if a partial import should be offered.
    const allMembersResponded = syncResult.totalRespondents === memberData.memberOrder.length;
    const isMembershipLocked = config.membershipLocked;
    const offerPartialImport = !(allMembersResponded && isMembershipLocked);

    // 5. Bundle all data and return it to the client.
    return {
      week: week,
      allCreatedWeeks: allCreatedWeeks || [],
      newMembers: formsData[week]?.newMembers || [],
      respondentIds: formsData[week]?.respondents || [],
      allMemberIds: memberData.memberOrder,
      members: memberData.members,
      gamePlanGames: gamePlan.games,
      startedGames: matchups,
      offerPartialImport: offerPartialImport
    };
  } catch (e) {
    console.error("Error in getFormImportData: ", e);
    // Re-throw the error so the client's onFailure handler gets it.
    throw new Error(`Failed to prepare for import: ${e.message}`);
  }
}

/**
 * [MODIFIED] This function now only launches the HTML file.
 */
function launchFormImport() {
  const html = HtmlService.createHtmlOutputFromFile('formImport')
      .setWidth(600)
      .setHeight(360);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Weekly Picks');
}

/**
 * [DEFINITIVE WORKER] Imports all processed picks into the correct weekly
 * pick'em sheet, survivor sheet, and eliminator sheet.
 *
 * @param {number} week The week to import.
 * @param {boolean} importOnlyStartedGames If true, only imports picks for games that have already started.
 */
function executePickImport(week, importOnlyStartedGames) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // --- 1. Fetch All Necessary Data ---
  
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration')) || {};
  const memberData = JSON.parse(docProps.getProperty('members')) || {};
  let formsData = JSON.parse(docProps.getProperty('forms')) || {};
  const databaseSheet = getDatabaseSheet();
  const responseSheet = databaseSheet.getSheetByName(`WK${week}`);

  // Parse the latest, de-duplicated picks from the response sheet.
  const parsedPicks = parseAllPicksFromSheet(responseSheet, memberData);
  // --- 2. Handle Pick'em Sheet Population ---
  if (config.pickemsInclude) {
    try {
      const weeklySheetName = `${weeklySheetPrefix}${week}`;
      let sheet = ss.getSheetByName(weeklySheetName);
      if (!sheet) {
        Logger.log(`🔁 No weekly sheet exists for week ${week}, creating one now...`);
        ss.toast(`Creating weekly sheet for week ${week}.`,'🔁 CREATING...');
        sheet = weeklySheet(ss,week,config,formsData,memberData,true);
      }
      // --- Create Lookup Maps ---
      // a) Member Name -> Row Index Map (Unchanged)
      const memberNameRange = ss.getRangeByName(`NAMES_${week}`);
      if (!memberNameRange) throw new Error(`Named range 'NAMES_${week}' not found.`);
      const memberNames = memberNameRange.getValues().flat();
      const memberNameToRowMap = new Map(memberNames.map((name, index) => [name, index]));

      // --- [THE NEW LOGIC] Team-Pair Matching ---
      // b) Matchup -> Column Index Map
      const matchupRange = ss.getRangeByName(`${LEAGUE}_${week}`);
      if (!matchupRange) throw new Error(`Named range '${LEAGUE}_${week}' not found.`);
      
      const matchupHeaders = matchupRange.getValues()[0];
      const matchupToColMap = new Map();
      matchupHeaders.forEach((header, index) => {
        const teams = header.toString().match(/[A-Z]{2,3}/g);
        if (teams && teams.length === 2) {
          const teamKey = teams.sort().join('-'); // e.g., "BUF-MIA"
          matchupToColMap.set(teamKey, index);
        }
      });
      // --- Prepare Data for Writing (Unchanged) ---
      const picksRange = ss.getRangeByName(`${LEAGUE}_PICKS_${week}`);
      const tiebreakerRange = ss.getRangeByName(`${LEAGUE}_TIEBREAKER_${week}`);
      let tiebreakers = tiebreakerRange.getValues();
      const commentRange = ss.getRangeByName(`COMMENTS_${week}`);
      let comments = commentRange.getValues();
      
      if (!picksRange) throw new Error(`Named range '${LEAGUE}_PICKS_${week}' not found.`);
      const picksData = picksRange.getValues();
      const gamePlan = formsData[week]?.gamePlan;
      let startedGames = new Set(getStartedGames());
  
      // --- 3. Loop Through Parsed Picks and Populate the 2D Array ---
      for (const memberId in parsedPicks) {
        const member = memberData.members[memberId];
        const picks = parsedPicks[memberId];
        const rowIndex = memberNameToRowMap.get(member.name);
        if (rowIndex === undefined) continue;
        
        for (const question in picks.pickem) {
          const pick = picks.pickem[question];

          // Find the original game from the gamePlan to get the team pair.
          const game = gamePlan.games.find(g => question.includes(g.awayTeamName) && question.includes(g.homeTeamName));
          if (game) {
            // Apply the import filter first for efficiency
            const matchupShortName = `${game.awayTeam} @ ${game.homeTeam}`;
            if (importOnlyStartedGames && !startedGames.has(matchupShortName)) {
              continue; // Skip if it's an upcoming game
            }
            
            // --- [THE NEW LOGIC] Find the column using the team-pair key ---
            const teamKey = [game.awayTeam, game.homeTeam].sort().join('-');
            const colIndex = matchupToColMap.get(teamKey);

            if (colIndex !== undefined) {
              // Check if the member's pick is actually one of the teams in the matchup
              if (pick === game.awayTeam || pick === game.homeTeam) {
                picksData[rowIndex][colIndex] = pick;
              }
            }
          }
        }
        if (picks.tiebreaker) tiebreakers[rowIndex][0] = picks.tiebreaker;
        if (picks.comments) comments[rowIndex][0] = picks.comments;
      }
      
      // --- 4. Write Data Back to the Sheet (Unchanged) ---
      picksRange.setValues(picksData);
      tiebreakerRange.setValues(tiebreakers);
      commentRange.setValues(comments);
      const text = `✅ Successfully imported Pick 'Em data into week '${week}' sheet.`;
      Logger.log(text);
      ss.toast(text,`PICK 'EMS IMPORT`)
    } catch (err) {
      const text = `❗ Failed to import Pick 'Em data into week '${week}' sheet.`;
      Logger.log(text + ' | ERROR: ' + err.stack);
      ss.toast(text,`PICK 'EMS FAILED`)

    }
  }

  // --- 5. Populate Survivor and Eliminator Sheets ---
  if (config.survivorInclude && week >= config.survivorStartWeek) {
    populateSurvElimSheet(ss, parsedPicks, memberData, config, formsData[week]?.gamePlan, week, 'survivor');
  }
  if (config.eliminatorInclude && week >= config.eliminatorStartWeek) {
    populateSurvElimSheet(ss, parsedPicks, memberData, config, formsData[week]?.gamePlan, week, 'eliminator');
  }
  
  // Updates the Outcomes sheet to reflect the games that were actually being evaluated by the form, resets conditional formatting and data validation rules, then checks if Pick 'Ems present, whether any values were in place on the Outcomes sheet already and replaces them after otherwise putting a connection in place back to the weekly sheet
  try {
    outcomesSheetUpdate(ss,week,config,formsData[week].gamePlan)
    const text = `✅ Successfully updated the OUTCOMES sheet input ranges for week '${week}' range.`;
    Logger.log(text);
    ss.toast(text,`OUTCOMES SHEET UPDATED`);
  } catch (err) {
    const text = `❗ Failed to update the OUTCOMES sheet input ranges for week '${week}' range.`;
    Logger.log(text + ' | ERROR: ' + err.stack);
    ss.toast(text,`OUTCOMES NOT UPDATED`);
  }

  // --- 6. Finalize and Save ---
  formsData[week].imported = true;
  saveProperties('forms', formsData);

  return { success: true, message: `✅ Picks for week ${week} have been successfully imported!` };
}

// You will also need this small helper function (modified from a previous version)
function getStartedGames() {
  try {
    const response = UrlFetchApp.fetch(SCOREBOARD);
    const data = JSON.parse(response.getContentText()).events;
    // An empty set is fine, as it is used to filter out games that are past kickoff and will not be imported
    const startedGames = new Set();
    for (const event of data) {
      if (event.status.type.state !== "pre") { // "in" or "post"
        startedGames.add(event.shortName);
      }
    }
    return startedGames;
  } catch (e) {
    return new Set();
  }
}

const filterMatchups = (matchupToColMap, startedGames) => {
  const startedTeams = new Set();
  
  // Extract all teams from started games
  startedGames.forEach(game => {
    const teams = game.toString().match(/([A-Z]{2,3})/g);
    if (teams) {
      teams.forEach(team => startedTeams.add(team));
    }
  });
  
  // Group by column index and filter
  const colGroups = {};
  Object.entries(matchupToColMap).forEach(([team, col]) => {
    if (!colGroups[col]) colGroups[col] = [];
    colGroups[col].push(team);
  });
  
  const result = {};
  Object.entries(colGroups).forEach(([col, teams]) => {
    if (teams.every(team => startedTeams.has(team))) {
      teams.forEach(team => result[team] = parseInt(col));
    }
  });
  
  return result;
};


/**
 * Reads a raw form response sheet, de-duplicates to "last
 * submission wins", and parses all pick types into a clean, structured object.
 * This is a read-only operation and does not modify any properties.
 *
 * @param {Sheet} sheet The Google Sheet object for a specific week's responses (e.g., 'WK1').
 * @param {Object} memberData The complete, current 'members' object.
 * @returns {Object} A "picks cache" object, where keys are member IDs and values
 *                   are objects containing all of that member's final picks.
 *                   e.g., { "id_123": { pickem: {...}, survivor: "BAL", ... } }
 */
function parseAllPicksFromSheet(sheet, memberData) {
  if (!sheet) {
    Logger.log("parseAllPicksFromSheet was called with a null sheet. Returning empty object.");
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log(`No responses found in sheet '${sheet.getName()}' to parse.`);
    return {};
  }

  const headers = data.shift();
  
  // --- Find critical column indexes using our robust helper ---
  const { nameCol, newNameCol } = findNameColumns(headers);
  if (nameCol === -1) {
    Logger.log("CRITICAL: Could not find a 'Select Your Name' column. Cannot parse picks.");
    return {};
  }

  // --- 1. De-duplicate to "Last Submission Wins" ---
  const latestSubmissions = {};
  const newUserAnswerRegex = /new user/i;
  data.forEach(row => {
    const name = newUserAnswerRegex.test(row[nameCol]) 
      ? row[newNameCol] 
      : row[nameCol];
    if (name && name.trim() !== '') {
      latestSubmissions[name.trim().toLowerCase()] = row;
    }
  });
  const finalResponseRows = Object.values(latestSubmissions);

  // --- 2. Create a lookup map for name -> ID ---
  const nameToIdMap = {};
  for (const id in memberData.members) {
    nameToIdMap[memberData.members[id].name.toLowerCase()] = id;
  }

  // --- 3. [THE NEW ENGINE] Intelligently Parse and Clean ALL Picks ---
  const weeklyPicksCache = {};
  // This regex will find the first 2-3 letter capital word in a string.
  const teamAbbrRegex = /[A-Z]{2,3}/; 
  const emojiAndSpecialCharsRegex = /[^A-Z0-9\s\-.()+]/gi;
  
  const survivorRegex = /survivor/i;
  const eliminatorRegex = /eliminator/i;
  const tiebreakerRegex = /tiebreaker/i;
  const commentsRegex = /comments/i;
  const pickemRegex = / at /i;

  finalResponseRows.forEach(row => {
    const name = newUserAnswerRegex.test(row[nameCol]) ? row[newNameCol] : row[nameCol];
    const memberId = nameToIdMap[name.trim().toLowerCase()];
    
    // If we can't map the submission to an existing member ID, we skip it.
    // (New members would have been added in the 'sync' step before this is called).
    if (!memberId) {
        Logger.log(`Skipping picks for "${name}" as they could not be mapped to a member ID.`);
        return; // 'continue' in a forEach loop
    }

    const userPicks = {
      pickem: {},
      survivor: null,
      eliminator: null,
      tiebreaker: null,
      comments: '' // Default comments to an empty string as requested
    };

    headers.forEach((header, index) => {
      let answer = row[index];
      if (answer === '' || answer === null || answer === undefined) return;
      
      // Convert to string for consistent processing
      answer = answer.toString().replace(emojiAndSpecialCharsRegex, '').trim();

      const question = header;

      if (survivorRegex.test(question) && answer) {
        userPicks.survivor = answer;
      } else if (eliminatorRegex.test(question) && answer) {
        userPicks.eliminator = answer;
      } else if (tiebreakerRegex.test(question)) {
        userPicks.tiebreaker = answer; // Tiebreaker is a number, not a team
      } else if (commentsRegex.test(question)) {
        userPicks.comments = answer;
      } else if (pickemRegex.test(question) && answer) {
        // The header is the matchup, the value is the cleaned team abbreviation
        userPicks.pickem[question] = answer;
      }
    });

    weeklyPicksCache[memberId] = userPicks;
  });

  return weeklyPicksCache;
}


/**
 * Fetches live scoreboard data to determine which games have ALREADY started.
 * @returns {Array<string>} An array of matchup short names, e.g., ["ARI @ LAR", "BUF @ MIA"].
 */
function getInvalidPickMatchups() {
  try {
    const response = UrlFetchApp.fetch(SCOREBOARD); // Your global SCOREBOARD constant
    const data = JSON.parse(response.getContentText()).events;
    const pastGames = [];
    
    for (const event of data) {
      // "pre" means the game has not started.
      if (event.status.type.state != "pre") {
        pastGames.push(event.shortName);
      }
    }
    return pastGames;
  } catch (err) {
    console.error("Could not fetch scoreboard data: " + err.toString());
    return []; // Return an empty array on failure
  }
}
















/**
 * Populates a contest sheet (Survivor or Eliminator) with the latest picks for a given week.
 *
 * @param {Sheet} ss The active Spreadsheet object.
 * @param {Object} parsedPicks The clean "picks cache" object from the parser.
 * @param {Object} memberData The complete 'members' object.
 * @param {Object} config The main 'configuration' object.
 * @param {Object} gamePlan The game plan for the week, containing spread data.
 * @param {number} week The week number to populate.
 * @param {string} contestType The type of sheet to populate: 'survivor' or 'eliminator'.
 */
function populateSurvElimSheet(ss, parsedPicks, memberData, config, gamePlan, week, contestType) {

  const sheetName = contestType.toUpperCase();
  try {
    let sheet = ss.getSheetByName(sheetName);

    // If the sheet doesn't exist, create it first.
    if (!sheet) {
      sheet = survElimSheet(ss, config, memberData, contestType);
    }

    const isAts = gamePlan[`${contestType}Ats`];
    
    // Widen the column for this week to accommodate spreads.
    const weekColumn = parseInt(week) + 4;
    if (isAts) sheet.setColumnWidth(weekColumn, 68);

    const nameToIdMap = new Map();
    if (memberData && memberData.members) {
      for (const id in memberData.members) {
        const member = memberData.members[id];
        if (member && member.name) {
          nameToIdMap.set(member.name.toString().trim().toLowerCase(), id);
        }
      }
    }

    // 2. Now, create the final map we need by reading the sheet: { id -> rowIndex }
    const memberIdToRowMap = new Map();
    ss.getRangeByName(`${sheetName}_NAMES`).getValues().flat().forEach((nameOnSheet, index) => {
      // Find the official ID for the name on the sheet using our case-insensitive map.
      const memberId = nameToIdMap.get(nameOnSheet.toString().trim().toLowerCase());
      if (memberId) {
        // If we found a match, create the link between the ID and its row on the sheet.
        memberIdToRowMap.set(memberId, index); // +2 for 0-based index and header row
      }
    });
    
    // --- The rest of your function can now work reliably --
    const dataRange = ss.getRangeByName(`${sheetName}_PICKS`);
    const writeRange = sheet.getRange(2,weekColumn,dataRange.getNumRows(),1)
    let writeArray = Array(dataRange.getNumRows()).fill(['']);
    for (const memberId in parsedPicks) {
      writeArray[memberIdToRowMap.get(memberId)][0] = parsedPicks[memberId]?.[contestType];
    }
    writeRange.setValues(writeArray);
    
    const text = `✅ Successfully populated ${sheetName} sheet for Week ${week}.`;
    Logger.log(text);
    ss.toast(text,`${sheetName} PICK IMPORT SUCCESS`);
    return true;
  } catch (err) {
    const text = `❗ Failed to populated ${sheetName} sheet for Week ${week}.`;
    Logger.log(text + '| ERROR: ' + err.stack);
    ss.toast(text,`${sheetName} PICK IMPORT FAILURE`);
  }
}

/**
 * [NEW & EFFICIENT] Updates an existing Survivor/Eliminator sheet with the latest
 * data for all members after an outcome has been processed.
 *
 * @param {Sheet} ss The active Spreadsheet object.
 * @param {string} contestType The type of sheet to update: 'survivor' or 'eliminator'.
 */
function updateSurvElimSheet(ss, contestType) {
  contestType = contestType.toLowerCase();
  
  // 1. Get the properties
  const docProps = PropertiesService.getDocumentProperties();
  let config = JSON.parse(docProps.getProperty('configuration')) || {};
  let memberData = JSON.parse(docProps.getProperty('members')) || {};

  const sheetName = contestType.toUpperCase();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const text = `❗ '${sheetName}' sheet not found, creating it now...`;
    Logger.log(text);
    ss.alert(text,`CREATING ${contestType.toUpperCase()} SHEET`);
    sheet = survElimSheet(ss,null,null,contestType);
  }

  // 2. Get the existing named ranges. This is much faster than rebuilding.
  const namesRange = ss.getRangeByName(`${sheetName}_NAMES`);
  const livesRange = ss.getRangeByName(`${sheetName}_LIVES`);
  const revivesRange = ss.getRangeByName(`${sheetName}_REVIVES`);
  const eliminatedRange = ss.getRangeByName(`${sheetName}_ELIMINATED`);
  const picksRange = ss.getRangeByName(`${sheetName}_PICKS`);
  
  if (!namesRange || !picksRange) {
    Logger.log(`Required named ranges for '${sheetName}' not found. Run the sheet builder.`);
    return;
  }

  const memberNamesOnSheet = namesRange.getValues().flat();
  const picksData = picksRange.getValues(); // Get the full 2D array of picks
  
  // Create a map of name -> row index for fast lookups.
  const nameToRowIndexMap = new Map(memberNamesOnSheet.map((name, index) => [name.toLowerCase(), index]));

  // --- 3. Prepare new data arrays for writing ---
  const newLivesData = [];
  const newRevivesData = [];
  const newEliminatedData = [];
  let done = true;
  // 4. Loop through the members AS THEY APPEAR ON THE SHEET.
  memberNamesOnSheet.forEach(name => {
    const member = Object.values(memberData.members).find(m => m.name.toLowerCase() === name.toLowerCase());
    
    if (member) {
      const livesKey = contestType === 'survivor' ? 'sL' : 'eL';
      const revivesKey = contestType === 'survivor' ? 'sR' : 'eR';
      const livesArray = member[livesKey] || [];
      const currentLives = livesArray.length > 0 ? livesArray[livesArray.length - 1] : (config[`${contestType}Lives`] || 1);
      const totalLives = parseInt(config[`${contestType}Lives`], 10) || 1;
      if (done && currentLives > 0) done = false;
      // a) Build the "Dots of Life" string.
      const livesDots = '🟢'.repeat(currentLives) + '⚫'.repeat(Math.max(0, totalLives - currentLives));
      newLivesData.push([livesDots]);

      // b) Get the revives count.
      newRevivesData.push([member[revivesKey] || 0]);

      // c) Get the elimination week.
      const elimWeek = member[`${contestType}EliminatedWeek`];
      newEliminatedData.push([elimWeek ? `Week ${elimWeek}` : '']);
      
      // d) Update the pick colors for this member.
      const rowIndex = nameToRowIndexMap.get(name.toLowerCase());
      const evalKey = contestType === 'survivor' ? 'sE' : 'eE';
      const evals = member[evalKey] || [];
      evals.forEach((isCorrect, weekIndex) => {
        const colIndex = weekIndex; // Assuming week columns start right after the fixed columns
        if (picksData[rowIndex][colIndex] !== '') {
          const cell = picksRange.getCell(rowIndex + 1, colIndex + 1);
          if (isCorrect === true) {
            cell.setBackground('#d9ead3').setFontLine(null);
          } else if (isCorrect === false) {
            cell.setBackground('#f4cccc').setFontLine('line-through');
          }
        }
      });

    } else {
      // If a member on the sheet isn't in our data, push blank values.
      newLivesData.push(['']);
      newRevivesData.push(['']);
      newEliminatedData.push(['']);
    }
  });

  // --- 5. Write the updated data to the sheet in a few, efficient calls ---
  livesRange.setValues(newLivesData);
  revivesRange.setValues(newRevivesData);
  eliminatedRange.setValues(newEliminatedData);
  if (config[`${contestType}Done`] != done) {
    config[`${contestType}Done`] = done;
    saveProperties('configuration',config);
  }
  Logger.log(`Successfully updated '${sheetName}' sheet.`);
}




/**
 * FUNCTIONS FOR TRIGGER AND UPDATE SURVIVOR/ELIMINATOR STATUS
 */


/**
 * [RUN ONCE] Creates the installable onEdit trigger for the spreadsheet.
 * The user should be instructed to run this function once from the script editor
 * to enable automatic score and status processing.
 */
function createOnEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggerFunctionName = 'onEditTrigger';

  // First, delete any existing triggers with the same function name to prevent duplicates.
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === triggerFunctionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create the new trigger to call our gatekeeper function.
  ScriptApp.newTrigger(triggerFunctionName)
    .forSpreadsheet(ss)
    .onEdit()
    .create();
    
  SpreadsheetApp.getUi().alert('Success!', 'The automatic score processing trigger has been installed. Statuses will now update automatically when game outcomes are entered.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Finds and deletes the specific installable onEdit trigger used for
 * automatic score processing ('onEditTrigger'). This provides a clean way
 * for an admin to disable the feature.
 */
function deleteOnEditTrigger() {
  let triggerDeleted = false;

  // 1. Get all triggers for the current project.
  const allTriggers = ScriptApp.getProjectTriggers();

  // 2. Loop through the triggers to find the specific one we want to delete.
  for (const trigger of allTriggers) {
    // We identify our trigger by the name of the function it is set to call.
    if (trigger.getHandlerFunction() === 'onEditTrigger') {
      
      // 3. If we find it, delete it.
      ScriptApp.deleteTrigger(trigger);
      triggerDeleted = true;
      
      // We break the loop because we assume there's only one.
      // Even if there were duplicates, this would safely delete the first one it finds.
      break; 
    }
  }

  // 4. Provide clear feedback to the user.
  if (triggerDeleted) {
    SpreadApp.getUi().alert('Success', 'The automatic score processing trigger has been successfully removed.', SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log("Automatic onEdit trigger was successfully deleted.");
  } else {
    SpreadApp.getUi().alert('Info', 'No automatic score processing trigger was found to delete.', SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log("No onEdit trigger was found to delete.");
  }
}

/**
 * This function is called by the installable onEdit trigger.
 * It efficiently checks if an edit was made to a relevant "outcome" cell on either
 * the main OUTCOMES sheet or a weekly sheet, then calls the main processing function.
 *
 * @param {Object} e The event object passed by the onEdit trigger.
 */
function onEditTrigger(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // --- Guard Clause 1: Was it a single cell edit? ---
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
    return;
  }

  let week = null;
  
  // --- Check 1: Was the edit on the main NFL_OUTCOMES sheet? ---
  if (sheetName === `${LEAGUE}_OUTCOMES`) {
    const editedRow = range.getRow();
    const editedCol = range.getColumn();
    // Check if the edit is in the data area (below headers) and in an odd-numbered (Winner) column.
    if (editedRow > 3) {
      week = Math.ceil(editedCol / 2);
    }
  } 
  // --- Check 2: Was the edit on a weekly sheet (e.g., "Week 3")? ---
  else if (sheetName.startsWith(weeklySheetPrefix)) {
    // Now, use a regex to be more precise and extract the week number.
    const weekRegex = new RegExp(`^${weeklySheetPrefix}(\\d{1,2})$`);
    const match = sheetName.match(weekRegex);
    
    if (match && match[1]) { // 'match[1]' contains the captured digits
      const potentialWeek = parseInt(match[1], 10);
      
      // Check if the edited cell is within one of the outcome named ranges
      const outcomeRange = ss.getRangeByName(`${LEAGUE}_PICKEMS_OUTCOMES_${potentialWeek}`);
      const marginRange = ss.getRangeByName(`${LEAGUE}_PICKEMS_OUTCOMES_${potentialWeek}_MARGIN`);
      
      // Check if the edited range is within either of these named ranges.
      // This is a more robust check than comparing A1 notations.
      if ( (outcomeRange && rangesOverlap(range, outcomeRange)) || (marginRange && rangesOverlap(range, marginRange)) ) {
        week = potentialWeek;
      }
    }
  }

  // --- Guard Clause 2: If we didn't identify a relevant week, exit. ---
  if (week === null) {
    return;
  }

  // --- Guard Clause 3: Final check to see if a form was even created for this week. ---
  const formsData = fetchProperties('forms');
  if (!formsData[week]) {
    Logger.log(`Edit detected for Week ${week}, but no form exists. Aborting process.`);
    return;
  }

  // --- If all checks pass, call the heavy lifter ---
  SpreadsheetApp.getActiveSpreadsheet().toast(`Change detected for Week ${week}. Processing scores...`, 'Updating', 5);
  
  try {
    evalSurvElimStatus(week, sheetName);
    SpreadsheetApp.getActiveSpreadsheet().toast(`✅ Pool Statuses for Week ${week} have been updated!`);
  } catch (err) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${err.message}`, '❌ Update Failed', 10);
    // Revert the change that triggered the error to signal failure to the user.
    if (e.oldValue !== undefined) {
      range.setValue(e.oldValue);
    }
  }
  
  function rangesOverlap(range1, range2) {
    return range1.getSheet().getName() === range2.getSheet().getName() &&
      range1.getLastRow() >= range2.getRow() &&
      range2.getLastRow() >= range1.getRow() &&
      range1.getLastColumn() >= range2.getColumn() &&
      range2.getLastColumn() >= range1.getColumn();
  }
}

/**
 * [THE DEFINITIVE HEAVY LIFTER] Evaluates and updates Survivor and Eliminator statuses for a given week.
 *
 * @param {number} week The week number to process.
 * @param {string} sourceSheetName The name of the sheet that triggered the edit.
 */
function evalSurvElimStatus(week, sourceSheetName) {
  const ss = fetchSpreadsheet();
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration')) || {};
  let memberData = JSON.parse(docProps.getProperty('members')) || {};
  const formsData = JSON.parse(docProps.getProperty('forms')) || {};
  
  const weeklySheetName = `${weeklySheetPrefix}${week}`;
  const weeklySheet = ss.getSheetByName(weeklySheetName);

  // --- Step 1: Sync Outcome Data (if triggered from a weekly sheet) ---
  if (weeklySheet && sourceSheetName === weeklySheetName) {
    const weeklyOutcomesRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
    const weeklyMarginsRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}_MARGIN`);
    const masterOutcomesRange = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`);
    const masterMarginsRange = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}_MARGIN`);

    // This is a simple, direct sync.
    masterOutcomesRange.setValues(weeklyOutcomesRange.getValues());
    masterMarginsRange.setValues(weeklyMarginsRange.getValues());
    Logger.log(`Synced outcomes from '${weeklySheet.getName()}' to the master OUTCOMES sheet.`);
  }

  // --- Step 2: Gather Official Outcomes ---
  const winners = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`).getValues().flat();
  const margins = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}_MARGIN`).getValues().flat();
  
  const gamePlan = formsData[week]?.gamePlan;
  if (!gamePlan || !gamePlan.games) {
    throw new Error(`Could not find a valid gamePlan with games for Week ${week}.`);
  }

  const matchups = gamePlan.games.map(g => `${g.awayTeam} @ ${g.homeTeam}`);
  
  const outcomeMap = new Map();
  gamePlan.games.forEach((game, index) => {
    const winner = winners[index];
    if (winner && winner.trim() !== '') {
      const matchupKey = `${game.awayTeam} @ ${game.homeTeam}`;
      const loser = (winner === 'TIE') ? 'TIE' : (winner === game.awayTeam ? game.homeTeam : game.awayTeam);
      
      outcomeMap.set(matchupKey, {
        winner: winner,
        loser: loser,
        margin: parseFloat(margins[index]) || 0,
        spread: game.spread // Carry the original spread along for ATS calcs
      });
    }
  });
  Logger.log(`Built Outcome Map for Week ${week}: Found ${outcomeMap.size} completed games.`);
  // --- Step 3: Process Survivor Sheet ---
  if (gamePlan.survivorInclude && week >= config.survivorStartWeek) {
    processContest(ss, week, 'SURVIVOR', memberData, outcomeMap, config);
  }

  // --- Step 4: Process Eliminator Sheet ---
  if (gamePlan.eliminatorInclude && week >= config.eliminatorStartWeek) {
    processContest(ss, week, 'ELIMINATOR', memberData, outcomeMap, config);
  }
  
  // --- Step 5: Save Updated Member Data ---
  saveProperties('members', memberData);

  // --- Step 6: (Optional) Check for contest end ---
  if (config.survivorInclude) {
    updateSurvElimSheet(ss, 'survivor');
  }
  if (config.eliminatorInclude) {
    updateSurvElimSheet(ss, 'eliminator');
  }
}

/**
 * [DEFINITIVE HELPER] A generic function to process either a Survivor or Eliminator contest.
 * This version uses a week-by-week array to track lives, enabling advanced history and revives.
 *
 * @param {Sheet} ss The active Spreadsheet object.
 * @param {number} week The week number being processed.
 * @param {string} contestType The type of sheet to populate: 'SURVIVOR' or 'ELIMINATOR'.
 * @param {Object} memberData The complete 'members' object.
 * @param {Object} outcomeMap A Map of the final game outcomes for the week.
 * @param {Object} config The main 'configuration' object.
 * @returns {Object} The modified and updated memberData object.
 */
function processContest(ss, week, contestType, memberData, outcomeMap, config) {
  const sheet = ss.getSheetByName(contestType.toUpperCase());
  if (!sheet) return memberData;

  const names = ss.getRangeByName(`${contestType}_NAMES`).getValues().flat();
  const nameToIdMap = new Map();
  for (const id in memberData.members) {
    if (memberData.members[id]?.name) {
      nameToIdMap.set(memberData.members[id].name.toLowerCase(), id);
    }
  }
  Logger.log(outcomeMap);
  // Define the keys we will use based on the contest type
  const picksKey = contestType === 'SURVIVOR' ? 'sP' : 'eP';
  const evalKey = contestType === 'SURVIVOR' ? 'sE' : 'eE';
  const isAts = config[`${contestType.toLowerCase()}Ats`] 
  const livesKey = contestType === 'SURVIVOR' ? 'sL' : 'eL';
  const livesSetting = config[`${contestType.toLowerCase()}Lives`] || 1;
  const startWeek = config[`${contestType.toLowerCase()}StartWeek`] || 1;
  const picks = ss.getRangeByName(`${contestType}_PICKS`).getValues();

  names.forEach((name, rowIndex) => {
    const memberId = nameToIdMap.get(name.trim().toLowerCase());
    if (!memberId) return;

    const member = memberData.members[memberId];
    const pickAbbr = picks[rowIndex][week - 1]?.toString().match(/[A-Z]{2,3}/)?.[0];
    if (!pickAbbr) return;
    
    // --- Update Pick History ---
    if (!member[picksKey]) member[picksKey] = [];
    member[picksKey][week - 1] = pickAbbr;

    // --- Determine Outcome ---
    const game = Array.from(outcomeMap.keys()).find(key => key.includes(pickAbbr));
    if (!game) return;

    const outcome = outcomeMap.get(game);
    let isCorrect = false;

    if (isAts) {
      isCorrect = calculateAtsResult(pickAbbr, outcome.winner, outcome.loser, outcome.margin, outcome.spread);
      if (contestType === 'ELIMINATOR') isCorrect = !isCorrect; // Flip logic for Eliminator
    } else { // Standard win/loss logic
      if (contestType === 'SURVIVOR') isCorrect = (outcome.winner === pickAbbr);
      if (contestType === 'ELIMINATOR') isCorrect = (outcome.loser === pickAbbr);
    }
    
    // --- 3. Initialize and Update Evaluation History (sE/eE) ---
    if (!member[evalKey]) member[evalKey] = [];
    member[evalKey][week - 1] = isCorrect;

    // --- 4. [THE NEW LOGIC] Calculate and Update Lives History (sL/eL) ---
    if (!member[livesKey]) member[livesKey] = [];
    
    // Determine the number of lives at the START of this week.
    let livesAtStartOfWeek;
    if (week == startWeek) {
        // If it's the first week of the contest, they start with the configured amount.
        livesAtStartOfWeek = parseInt(livesSetting, 10);
    } else {
        // Otherwise, they start with the number of lives they had at the END of the previous week.
        // The `|| 0` handles cases where they might have missed a week.
        livesAtStartOfWeek = member[livesKey][week - 2] || 0;
    }
    
    // Calculate lives at the END of this week.
    let livesAtEndOfWeek = livesAtStartOfWeek;
    if (!isCorrect && livesAtStartOfWeek > 0) {
      livesAtEndOfWeek--; // Decrement a life for an incorrect pick.
    }
    
    // Save the final life count for this week.
    member[livesKey][week - 1] = livesAtEndOfWeek;
    
    // Update the elimination week if they just ran out of lives.
    if (livesAtEndOfWeek === 0 && livesAtStartOfWeek > 0) {
        member[`${contestType.toLowerCase().substring(0,1)}O`] = week; // Out week
    }
  });

  return memberData;
}

/**
 * Calculates if a pick was correct against the spread,
 * using only the winning team and the margin of victory.
 *
 * @param {string} pick The member's picked team abbreviation (e.g., "DAL").
 * @param {string} winner The actual winning team's abbreviation (e.g., "PHI").
 * @param {number} margin The positive margin of victory (e.g., 3).
 * @param {string} spread The original spread string from the game plan (e.g., "DAL -6.5").
 * @returns {boolean} True if the pick was correct against the spread.
 */
function calculateAtsResult(pick, winner, loser, margin, spread) {
  if (!spread || spread.toUpperCase() === 'PK' || spread.trim() === '0') {
    // If it's a "pick'em", the pick is correct if they picked the winner.
    return pick === winner;
  }

  try {
    const spreadMatch = spread.match(/([A-Z]{2,3})\s*([+-]?\d+\.?\d*)/);
    if (!spreadMatch) return false; // Invalid spread format

    const [, favoriteTeam, spreadValueStr] = spreadMatch;
    const spreadValue = parseFloat(spreadValueStr); // e.g., -6.5
    const underdogTeam = favoriteTeam == winner ? loser : winner;
    if (pick === underdogTeam) {
      return true;
    } else if (pick === favoriteTeam && margin > Math.abs(spreadValue)) {
      // They picked the FAVORITE. They win if the actual winner is the favorite AND the margin is greater than the spread.
      // e.g., Spread is -6.5, margin must be 7 or more.
      return true;
    } else {
      // There was a TIE or the favorite won by less than the spread value
      return false;
    }
  } catch (e) {
    console.error(`Error calculating ATS result for spread "${spread}": ${e.toString()}`);
    return false;
  }
}

























































/**
 * Wrapper function to call the sync process from the spreadsheet menu.
 */
function syncCurrentWeekResponses() {
  const ui = SpreadsheetApp.getUi();
  try {
    const week = fetchWeek(false, true);
    // Show a toast to indicate the process is starting.
    SpreadsheetApp.getActiveSpreadsheet().toast('Syncing responses for Week ' + week + '...', 'In Progress', 10);
    const result = syncFormResponses(week);
    ui.alert('Sync Complete', `${result.message}\nFound ${result.newMembers} new member(s).\nTotal Respondents: ${result.totalRespondents}.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Error', `Failed to sync responses: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * [DEFINITIVE CONTROLLER] Orchestrates the entire process of fetching, de-duplicating,
 * and processing form responses to update the 'members' and 'forms' properties.
 */
function syncFormResponses(week) {
  // --- 1. SETUP ---
  Logger.log('Getting things ready to import for week ' + week);
  week = week || fetchWeek();
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration'));
  let memberData = JSON.parse(docProps.getProperty('members'));
  let formsData = JSON.parse(docProps.getProperty('forms'));
  const responseSheet = getDatabaseSheet().getSheetByName(`WK${week}`);
  
  if (!responseSheet) {
    Logger.log(`No response sheet found for WK${week}.`);
    return { success: true, message: `No response sheet exists for Week ${week}.`, newMembers: 0, totalRespondents: 0 };
  }
  
  const data = responseSheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log(`No responses found in sheet for WK${week}.`);
    return { success: true, message: `No responses to sync for Week ${week}.`, newMembers: 0, totalRespondents: 0 };
  }

  const headers = data.shift();
  const { nameCol, newNameCol } = findNameColumns(headers);
  const newUserAnswerRegex = /new user/i;
  
  // --- 2. DE-DUPLICATE RESPONSES ("Last Submission Wins") ---
  const latestSubmissions = {};
  data.forEach(row => {
    const name = newUserAnswerRegex.test(row[nameCol]) ? row[newNameCol] : row[nameCol];
    if (name && name.trim() !== '') {
      latestSubmissions[name.trim().toLowerCase()] = row;
    }
  });
  const finalResponseRows = Object.values(latestSubmissions);

  // --- 3. PROCESS NEW MEMBERS ---
  const nameToIdMap = {};
  for (const id in memberData.members) {
    nameToIdMap[memberData.members[id].name.toLowerCase()] = id;
  }
  const newMemberIds = [];
  
  finalResponseRows.forEach(row => {
    const submitterChoice = row[nameCol];
    const newUserName = (newNameCol > -1) ? row[newNameCol].trim() : '';
    
    if (newUserAnswerRegex.test(submitterChoice) && newUserName) {
      const nameKey = newUserName.trim().toLowerCase();
      // If their name is NOT in our official map, they are a new member.
      if (!nameToIdMap[nameKey]) {
        const permanentId = generateUniqueId();
        
        memberData.memberOrder.push(permanentId);
        memberData.members[permanentId] = createNewMember(
          newUserName,
          false, // New members from the form are always initially unpaid
          week // The join week is the week of the form they submitted
        );
        
        nameToIdMap[nameKey] = permanentId;
        newMemberIds.push(permanentId);
      }
    }
  });

  // --- 4. UPDATE RESPONDENT LIST IN 'forms' OBJECT ---
  const respondentIds = finalResponseRows.map(row => {
    const submitterChoice = row[nameCol];
    const newUserName = (newNameCol > -1) ? row[newNameCol] : '';
    const name = newUserAnswerRegex.test(submitterChoice) ? newUserName : submitterChoice;
    return nameToIdMap[name.trim().toLowerCase()];
  }).filter(id => id);
  
  if (formsData[week]) {
    formsData[week].respondents = [...new Set(respondentIds)];
    if (newMemberIds.length > 0) {
      if (formsData[week].hasOwnProperty('newMembers')) {
        formsData[week].newMembers.push(...newMemberIds);
      } else {
        formsData[week].newMembers = newMemberIds;
      }
    }
    formsData[week].responseCount = formsData[week].respondents.length;
    formsData[week].lastResponseTime = new Date().toISOString();
    const allMemberIds = memberData.memberOrder;
    formsData[week].nonRespondents = allMemberIds.filter(id => !respondentIds.includes(id));
  }

  formsData[week].imported = false;

  // --- 5. SAVE UPDATED DATA ---
  saveProperties('members', memberData);
  saveProperties('forms', formsData);

  // const newMembers = newMemberIds.map(id => memberData.members[id]?.name);
  return {
    success: true,
    message: `Sync complete for Week ${week}.`,
    newMembers: formsData[week].newMembers,
    totalRespondents: respondentIds.length
  };
}

/**
 * Finds the column indexes for name fields using regex.
 */
function findNameColumns(headers) {
  const mainNameRegex = /select.*name/i; 
  const newUserNameRegex = /enter.*name/i;
  let nameCol = -1, newNameCol = -1;

  headers.forEach((header, index) => {
    if (mainNameRegex.test(header)) nameCol = index;
    else if (newUserNameRegex.test(header)) newNameCol = index;
  });
  
  return { nameCol, newNameCol };
}

/** 
 * Function to eliminate the "New User" entry on the specified week's form
 */
function removeNewUserQuestion(week) {
  let nameQuestion, found = false;
  try {
    const id = fetchProperties('forms')[week].id;
    let form = FormApp.openById(id);
    let items = form.getItems();
    for (let a = 0; a < items.length; a++) {
      if (items[a].getType() == 'LIST' && items[a].getTitle() == 'Name') {
        nameQuestion = items[a];
      }
    }

    let choices = nameQuestion.asListItem().getChoices();
    for (let a = 0; a < choices.length; a++) {
      if (choices[a].getValue() == 'New User') {
        choices.splice(a,1);
        found = true;
      }
    }
    if (found) {
      nameQuestion.asListItem().setChoices(choices);
      ss.toast('Removed the option of \"New User\" from the form.');
    } else {
      ss.toast('No \"New User\" option was present on the form.');
    }
  }
  catch (err) {
    ss.toast('Failed to remove the list item of \"New User\" from the form.');
  }
}

// ============================================================================================================================================
// UTILITIES
// ============================================================================================================================================

/**
 * Displays a clean modal dialog with a link for the user to click.
 * This is the standard way to direct a user to a URL from a server-side script.
 *
 * @param {string} url The URL the link should point to.
 * @param {string} title The title for the dialog window.
 * @param {string} linkText The text to display for the link itself.
 */
function showLinkDialog(url, title, linkText, subText) {
  const htmlContent = `
    <div style="font-family: 'Montserrat', sans-serif; text-align: center; padding: 20px;">
      <p style="font-size: 16px;">
        <a href="${url}" target="_blank" onclick="google.script.host.close()" 
           style="font-weight: bold; text-decoration: none; background-color: #013369; color: white; padding: 10px 20px; border-radius: 5px;">
          ${linkText}
        </a>
      </p>
      ${subText ? '<div style="font-size: 12px;">' + subText + '</div>' : ''}
    </div>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent).setWidth(400).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

/**
 * Generates a simple, robust, and sufficiently unique ID string.
 * Creates an ID like "id_1234567890".
 *
 * @returns {string} A new unique ID.
 */
function generateUniqueId() {
  const randomPart = Math.random().toString(36).substring(3, 15).toUpperCase();
  return `id_${randomPart}`;
}

// RESET Function to reset and create menu for runFirst
function resetSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Return to spreadsheet for prompts');
  let prompt = ui.alert('Reset spreadsheet and delete all data?', ui.ButtonSet.YES_NO);
  if (prompt == 'YES') {
    
    let promptTwo = ui.alert('Are you sure? This would be very difficult to recover from.',ui.ButtonSet.YES_NO);
    if (promptTwo == 'YES') {
      let ranges = ss.getNamedRanges();
      for (let a = 0; a < ranges.length; a++){
        ranges[a].remove();
      }
      let sheets = ss.getSheets();
      let baseSheet = ss.insertSheet();
      for (let a = 0; a < sheets.length; a++){
        ss.deleteSheet(sheets[a]);
      }
      let protections = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (let a = 0; a < protections.length; a++){
        protections[a].remove();
      }
      protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (let a = 0; a < protections.length; a++){
        protections[a].remove();
      }      
      baseSheet.setName('Sheet1');

      // Deletes initialization, time zone, and any other response-associated properites
      let properties = PropertiesService.getDocumentProperties();
      properties.deleteAllProperties();

      deleteTriggers();

      initializeMenu();

    } else {
      ss.toast('Canceled reset');
    }
  } else {
    ss.toast('Canceled reset');
  }
  
}

// FETCH SPREADSHEET - Checks that the 'ss' variable passed into a script is not null, undefined, or a non-spreadsheet
function fetchSpreadsheet(ss) {
  try {
    if (ss && typeof ss.getSheets === 'function' && typeof ss.getId === 'function') {
      return ss;
    } else {
      throw new Error('Invalid Spreadsheet object');
    }
  } catch (err) {
    if (ss !== null && ss !== undefined) {
      Logger.log('ALERT: The function \'' + (new Error()).stack.split('\n')[2].trim().split(' ')[1] + '\' passed ' + typeof ss + ' \'' + ss + '\' to the \'fetchSpreadsheet\' function.');
      Logger.log(err.stack);
    }
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  return ss;
}

// FETCH UI - Checks that the 'ui' variable passed into a script is not null, undefined, or a non-UI
function fetchUi(ui) {
  try{
    if (typeof ui.showModalDialog !== 'function') {
      throw new Error('Non-UI passed');
    }
  }
  catch (err) {
    if (ui !== null && ui !== undefined) {
      Logger.log('ALERT: The function \'' + (new Error()).stack.split('\n')[2].trim().split(' ')[1] + '\' passed ' + typeof ui + ' \'' + ui + '\' to the \'fetchUi\' function.');
    }
    ui = SpreadsheetApp.getUi();
  }
  return ui;
}

// SERVICE Function to remove all triggers on project
function deleteTriggers() {
  let triggers = ScriptApp.getProjectTriggers();
  for (let a = 0; a < triggers.length; a++) {
    ScriptApp.deleteTrigger(triggers[a]);
  }
}

// ADJUST ROWS - Cleans up rows of a sheet by providing the total rows that currently exist with data
function adjustRows(sheet,rows,verbose){
  let maxRows = sheet.getMaxRows(); 
  if (rows == undefined || rows == null) {
    rows = sheet.getLastRow();
  }
  if (rows > 0 && rows > maxRows) {
    sheet.insertRowsAfter(maxRows,(rows-maxRows));
    if(verbose) return Logger.log('Added ' + (rows-maxRows) + ' rows');
  } else if (rows < maxRows && rows != 0){
    sheet.deleteRows((rows+1), (maxRows-rows));
    if(verbose) return Logger.log('Removed ' + (maxRows - rows) + ' rows');
  } else {
    if(verbose) return Logger.log('Rows not adjusted');
  }
}

// ADJUST COLUMNS - Cleans up columns of a sheet by providing the total columns that currently exist with data
function adjustColumns(sheet,columns,verbose){
  let maxColumns = sheet.getMaxColumns(); 
  if (columns == undefined || columns == null) {
    columns = sheet.getLastColumn();
  }
  if (columns > 0 && columns > maxColumns) {
    sheet.insertColumnsAfter(maxColumns,(columns-maxColumns));
    if(verbose) return Logger.log('Added ' + (columns-maxColumns) + ' columns');
  }  else if (columns < maxColumns && columns != 0){
    sheet.deleteColumns((columns+1), (maxColumns-columns));
    if(verbose) return Logger.log('Removed ' + (maxColumns - columns) + ' column(s)');
  } else {
    if(verbose) return Logger.log('Columns not adjusted');
  }
}

// GENERATES HEX GRADIENT - Provide a start and end and a count of values and this function generates a HEX gradient. Midpoint value is optional.
function hexGradient(start, end, count, midpoint) { // start and end in either 3 or 6 digit hex values, count is total values in array to return
  if (count < 2 || count.isNaN) {
    Logger.log('ERROR: Please provide a \'count\' value of 2 or greater');
    return null;
  } else {
    count = Math.ceil(count);
    if (midpoint == null || midpoint == undefined) {
      // strip the leading # if it's there
      start = start.replace(/^\s*#|\s*$/g, '');
      end = end.replace(/^\s*#|\s*$/g, '');

      // convert 3 char codes --> 6, e.g. `E0F` --> `EE00FF`
      if(start.length == 3){
        start = start.replace(/(.)/g, '$1$1');
      }

      if(end.length == 3){
        end = end.replace(/(.)/g, '$1$1');
      }

      let arr = ['#'+start];
      let tmpRed, tmpGreen, tmpBlue;

      // get colors
      let startRed = parseInt(start.substr(0, 2), 16),
          startGreen = parseInt(start.substr(2, 2), 16),
          startBlue = parseInt(start.substr(4, 2), 16);
      let endRed = parseInt(end.substr(0, 2), 16),
          endGreen = parseInt(end.substr(2, 2), 16),
          endBlue = parseInt(end.substr(4, 2), 16);
      let stepRed = (endRed-startRed)/(count-1),
          stepGreen = (endGreen-startGreen)/(count-1),
          stepBlue = (endBlue-startBlue)/(count-1);
      
      for (let a = 1; a < count-1; a++) {
        // calculate the step differential for each color
        tmpRed = ((stepRed * a) + startRed).toString(16).split('.')[0];
        tmpGreen = ((stepGreen * a) + startGreen).toString(16).split('.')[0];
        tmpBlue = ((stepBlue * a) + startBlue).toString(16).split('.')[0];
        // ensure 2 digits by color
        if( tmpRed.length == 1 ) tmpRed = '0' + tmpRed;
        if( tmpGreen.length == 1 ) tmpGreen = '0' + tmpGreen;
        if( tmpBlue.length == 1 ) tmpBlue = '0' + tmpBlue;
        arr.push(('#' + tmpRed + tmpGreen + tmpBlue).toUpperCase());
      }
      arr.push('#'+end);
      return arr;
    } else {
      count = Math.ceil(count);
      if (count % 2 == 0) {
        count++;
        // Logger.log('Even number provided with midpoint, increasing count to ' + count);
      }
      let half = Math.ceil(count/2);
      let arr = hexGradient(start,midpoint,half);
      arr.pop();
      let arr2 = hexGradient(midpoint,end,half);
      arr = arr.concat(arr2);
      return arr;
    }
  }
}

// ENSURE ARRAY IS RECTANGULAR - a function to ensure that an array has blank values if it fails to have a full set of columns per row
function makeArrayRectangular(arr) {
  const maxLength = Math.max(...arr.map(row => row.length));
  for (let a = 0; a < arr.length; a++) {
    // While the row's length is less than the maximum length, push a placeholder value
    while (arr[a].length < maxLength) {
      arr[a].push('');
    }
  }
  return arr;
}

// GET TIMEZONE
function timezoneSet() {
  // Get the value for the script property timezone
  const scriptProperties = PropertiesService.getDocumentProperties();
  const tz = scriptProperties.getProperty('tz');
  if (tz != null) {
    return true;
  } else {
    Logger.log('No timezone confirmation has been done yet');
    return false;
  }
}

// SET PROPRTY - sets a script property based on an inputted name (string) and a value (string/array/object) (essentially this ia global variable)
function setProperty(property,value){
  const scriptProperties = PropertiesService.getDocumentProperties();
  if (typeof value === 'object' && !Array.isArray(value) && value !== null) {
    scriptProperties.setProperty(property,JSON.stringify(value));
  } else {
    scriptProperties.setProperty(property,value);
  }
}

// OPEN URL - Quick script to open a new tab with the newly created form, in this case
function openUrl(url,week){
  if (!url || typeof url !== 'string') {
    throw new Error("Invalid URL provided.");
  }
  if (week == null) {
    week = SpreadsheetApp.getActiveSpreadsheet().getRangeByName('WEEK').getValue();
  }
  if (week == undefined) {
    week = fetchWeek();
  }

  // Create the HTML content with the Montserrat font
  let htmlContent = `
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap" rel="stylesheet">
    <div style="font-family: 'Montserrat', sans-serif; text-align: center; padding: 20px;">
      <p style="font-size: 22px;"><a href="${url}" target="_blank" style="font-weight: bold;">Click for Week ` + week + ` Form</a></p>
    </div>
  `;

  let htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(350)
    .setHeight(180);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

// VIEW USER PROPERTIES - Shows all set variables within Google user properties
// This is a back-end and unused script, these variables aren't isolated to the sheet/script but used by the form/sheet connection when triggering onSubmit calls
function viewUserProperties() {
  let userProperties = PropertiesService.getUserProperties().getProperties();
  Logger.log('User Properties:');
  for (let key in userProperties) {
    Logger.log(key + ': ' + userProperties[key]);
  }
}

// VIEW SCRIPT PROPERTIES - Shows all set variables within Google user properties
// This is a back-end and unused script, these variables aren't isolated to the sheet/script but used by the form/sheet connection when triggering onSubmit calls
function viewScriptProperties() {
  let scriptProperties = PropertiesService.getScriptProperties().getProperties();
  Logger.log('Script Properties:');
  for (let key in scriptProperties) {
    Logger.log(key + ': ' + scriptProperties[key]);
  }
}

// VIEW DOCUMENT PROPERTIES - Shows all set variables within Google user properties
// This is a back-end and unused script, these variables aren't isolated to the sheet/script but used by the form/sheet connection when triggering onSubmit calls
function viewDocumentProperties() {
  let documentProperties = PropertiesService.getDocumentProperties().getProperties();
  Logger.log('Document Properties:');
  for (let key in documentProperties) {
    Logger.log(key + ': ' + documentProperties[key]);
  }
}


/**
 * A simple toast message helper.
 */
function showToast(message,ss) {
  ss = fetchSpreadsheet(ss);
  ss.toast(message);
}


// ============================================================================================================================================
// CORE SHEETS
// ============================================================================================================================================

/**
 * Creates or updates the OUTCOMES sheet with a multi-row header, 
 * custom color scales for margins, and robust handling of empty playoff weeks.
 */
function outcomesSheet(ss) {
  ss = fetchSpreadsheet(ss);
  const sheetName = `${LEAGUE}_OUTCOMES`;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  // --- Start with a clean slate ---
  sheet.clear();
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).clearDataValidations().clearNote();
  sheet.getNamedRanges().forEach(namedRange => {
    if (namedRange.getName().startsWith(`${LEAGUE}_OUTCOMES`)) {
      namedRange.remove();
    }
  });
  sheet.setTabColor(dayColorsFilled[dayColorsFilled.length - 1]); // Your custom tab color

  const data = ss.getRangeByName(LEAGUE)?.getValues() || fetchSchedule(ss);
  
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  // --- 1. Build the New, 3-Row Header Structure ---
  const weekTypeHeaders = []; // For Row 2 (e.g., "Regular Season", "WildCard")
  const weekNumHeaders = [];  // For Row 3 (e.g., "Week 1", "Week 2")

  for (const a in weeks) {
    const weekName = WEEKNAME[weeks[a]] ? WEEKNAME[weeks[a]].name : "Regular Season";
    weekTypeHeaders.push(weekName, ""); // Add name and a blank for the margin column
    weekNumHeaders.push(`Week ${weeks[a]}`, ""); // Add week # and a blank for the margin column|
  }
  
  // --- 2. Sheet Resizing and Basic Formatting ---
  const headerRow1 = 1; // Main Title
  const headerRow2 = 2; // Week Type (e.g., WildCard)
  const headerRow3 = 3; // Week Number
  const dataStartRow = 4; // Data starts on row 4 now

  const totalCols = weekNumHeaders.length;
  const totalRows = dataStartRow + MAXGAMES -1;
  
  if (sheet.getMaxColumns() < totalCols) sheet.insertColumnsAfter(1, totalCols - 1);
  if (sheet.getMaxRows() < totalRows) sheet.insertRowsAfter(1, totalRows - 1);
  if (sheet.getMaxColumns() > totalCols) sheet.deleteColumns(totalCols+1,sheet.getMaxColumns()-totalCols);
  if (sheet.getMaxRows() > totalRows) sheet.deleteRows(totalRows+1,sheet.getMaxRows()-totalRows);
  
  sheet.getRange(headerRow3+1,1,totalRows-headerRow3,weekNumHeaders.length).setBackground('#dddddd');

  // --- 3. Apply Header and Base Styles ---
  // Main Title (Row 1)
  sheet.getRange(headerRow1, 1, 1, totalCols).mergeAcross().setValue(sheetName.replace(/_/g, ' '))
      .setFontWeight('bold').setFontSize(18).setFontFamily("Montserrat")
      .setBackground('#666666').setFontColor('#ffffff')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(headerRow1, 40);

  // Week Type Headers (Row 2)
  sheet.getRange(headerRow2, 1, 1, totalCols).setValues([weekTypeHeaders])
      .setBackground('#333333').setFontColor('#ffffff').setFontWeight('bold').setFontStyle('italic').setFontSize(8)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(headerRow2, 25);
  
  // Week Number Headers (Row 3)
  sheet.getRange(headerRow3, 1, 1, totalCols).setValues([weekNumHeaders])
      .setBackground('#000000').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(headerRow3, 25);
  
  // Merge the header cells in pairs
  for (let i = 1; i <= totalCols; i += 2) {
      sheet.getRange(headerRow2, i, 1, 2).mergeAcross();
      sheet.getRange(headerRow3, i, 1, 2).mergeAcross();
  }

  const dataBodyRange = sheet.getRange(dataStartRow, 1, MAXGAMES, totalCols);
  dataBodyRange.setFontFamily("Montserrat").setFontSize(9).setVerticalAlignment('middle').setHorizontalAlignment('center');

  // --- 4. Populate Matchups, Set Ranges, Validation, and Formatting ---
  let allConditionalFormatRules = [];
  const marginValidationRule = SpreadsheetApp.newDataValidation().requireNumberBetween(0, 50).build();
  let col = 1;
  for (const a in weeks) {
    const winnerCol = (col - 1) * 2 + 1;
    const marginCol = winnerCol + 1;
    sheet.setColumnWidth(winnerCol, 50);
    sheet.setColumnWidth(marginCol, 50);

    // [THE FIX] Filter for the week's matchups, gracefully handling empty weeks.
    const weekMatchups = data.filter(row => row[0] == weeks[a]);
    
    // If no matchups for this week (e.g., future playoff week), skip to the next loop iteration.
    if (weekMatchups.length === 0) {
      continue; 
    }

    const winnerRange = sheet.getRange(dataStartRow, winnerCol, weekMatchups.length);
    ss.setNamedRange(`${LEAGUE}_OUTCOMES_${weeks[a]}`, winnerRange);
    const marginRange = sheet.getRange(dataStartRow, marginCol, weekMatchups.length);
    ss.setNamedRange(`${LEAGUE}_OUTCOMES_${weeks[a]}_MARGIN`, marginRange);
    marginRange.setDataValidation(marginValidationRule);

    weekMatchups.forEach((game, index) => {
      const rowIndex = dataStartRow + index;
      const winnerCell = sheet.getRange(rowIndex, winnerCol);
      const marginCell = sheet.getRange(rowIndex, marginCol);
      const awayTeam = game[6];
      const homeTeam = game[7];
      const dayIndex = game[2] + 3; // Numeric day used for gradient application (-3 is Thursday, 1 is Monday);
      
      winnerCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([awayTeam, homeTeam, 'TIE'], true).setAllowInvalid(false).build());
      marginCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from({ length: 46 }, (_, index) => index), true).build());

      // --- [THE NEW COLOR LOGIC] ---
      // Set the base background color for BOTH cells based on the day of the week.
      winnerCell.setBackground(dayColors[dayIndex]);
      marginCell.setBackground(dayColors[dayIndex]);

      // Create a conditional format rule for the winning team's background color.
      // This will override the base color only when a winner is selected.
      const homeWinRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(homeTeam)
        .setBold(true)
        .setBackground(dayColorsFilled[dayIndex])
        .setRanges([winnerCell])
        .build();
      allConditionalFormatRules.push(homeWinRule);
      
      const awayWinRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(awayTeam)
        .setBold(false) // Home team is bold, away is not
        .setBackground(dayColorsFilled[dayIndex])
        .setRanges([winnerCell])
        .build();
      allConditionalFormatRules.push(awayWinRule);

      // [BONUS] Custom color scale for the margin, using the day's colors.
      // It will scale from the lighter day color to the darker filled day color.
      const marginColorScaleRule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue(dayColors[dayIndex], SpreadsheetApp.InterpolationType.NUMBER, '1')
        .setGradientMaxpointWithValue(dayColorsFilled[dayIndex], SpreadsheetApp.InterpolationType.NUMBER, '10')
        .setRanges([marginCell])
        .build();
      allConditionalFormatRules.push(marginColorScaleRule);
    });
    col++;
  }
  
  // Add a single rule for TIEs for all columns
  const tieRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TIE')
    .setBold(false)
    .setBackground('#aaaaaa')
    .setRanges([sheet.getRange(dataStartRow,1,sheet.getMaxRows()-dataStartRow,sheet.getLastColumn())])
    .build();
  allConditionalFormatRules.unshift(tieRule);
  // Add a single rule for TIEs for all columns
  const zeroRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBold(false)
    .setBackground('#aaaaaa')
    .setRanges([sheet.getRange(dataStartRow,1,sheet.getMaxRows()-dataStartRow,sheet.getLastColumn())])
    .build();
  allConditionalFormatRules.unshift(zeroRule);
  sheet.setConditionalFormatRules(allConditionalFormatRules);

  Logger.log(`Completed setting up ${LEAGUE} OUTCOMES sheet`);
}


/** 
 * UPDATE OUTCOMES - Updates the data validation, color scheme, and matchups for a specific week on the winners sheet
 * 
 * Modified to ensure that the matchups displayed are parallel to those from the weekly sheet -- only called by the "weeklySheet" function
 * 
*/
function outcomesSheetUpdate(ss,week,config,gamePlan) {
  ss = ss || fetchSpreadsheet(ss);
  week = week || fetchWeek();
  let sheet = ss.getSheetByName(`${LEAGUE}_OUTCOMES`);
  let matchups = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`);
  let margins = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}_MARGIN`);
  if (!sheet || !matchups) {
    const missingOutcomes = `⚠️ Outcomes sheet error or not present, creating now`;
    Logger.log(missingOutcomes);
    ss.toast('OUTCOMES SHEET ISSUE',missingOutcomes);
    sheet = outcomesSheet(ss);
    matchups = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}`);
    margins = ss.getRangeByName(`${LEAGUE}_OUTCOMES_${week}_MARGIN`);
  }
  const startRow = matchups.getRow(); // First row of matchups on 

  let docProps;
  if (!config || !gamePlan) {
    docProps = PropertiesService.getDocumentProperties();
  }
  gamePlan = gamePlan || JSON.parse(docProps.getProperty('forms'))[week].gamePlan || {};
  config = config || JSON.parse(docProps.getProperty('configuration')) || {};

  const contests = gamePlan.games;
  
  // Clears data validation and notes
  
  matchups.clearDataValidations().clearNote();
  margins.clearDataValidations().clearNote();
  
  let existingRules = sheet.getConditionalFormatRules();
  let rulesToKeep = [];
  let newRules = [];
  for (let a = 0; a < existingRules.length; a++) {
    let ranges = existingRules[a].getRanges();
    for (let b = 0; b < ranges.length; b++) {
      if (ranges[b].getColumn() != matchups.getColumn() && ranges[b].getColumn() != margins.getColumn()) {
        rulesToKeep.push(existingRules[a]);
      }
    }
  }
  sheet.clearConditionalFormatRules();
  
  sheet.getRange(startRow,matchups.getColumn(),sheet.getMaxRows()-startRow+1,1).setBackground('#dddddd');
  sheet.getRange(startRow,margins.getColumn(),sheet.getMaxRows()-startRow+1,1).setBackground('#dddddd');
  let start = startRow;
  let end = start+1;
  let teams = []; // Array to cross-reference existing values when re-writing
  
  for (let a = 0; a < contests.length; a++) {
    teams.push(contests[a].awayTeam);
    teams.push(contests[a].homeTeam);
    sheet.getRange(a+startRow,matchups.getColumn()).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([contests[a].awayTeam,contests[a].homeTeam,'TIE'], true).build());
    // Color Coding Days
    if (contests[a].dayName != contests[a+1]?.dayName) {
      // Matchup column color (static) and conditional formatting
      matchupCell = sheet.getRange(start,matchups.getColumn(),end-start,1)
      matchupCell.setBackground(dayColorsObj[contests[a].dayName]);
      let homeWin = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=iferror(match(indirect("R[0]C[0]",false),indirect("${LEAGUE}_HOME_${week}"),0)>=0,false)`)
        .setBackground(dayColorsFilledObj[contests[a].dayName])
        .setBold(true)
        .setRanges([matchupCell])
        .build();
      newRules.push(homeWin);
      let awayWin = SpreadsheetApp.newConditionalFormatRule()
        .whenCellNotEmpty()
        .setBackground(dayColorsFilledObj[contests[a].dayName])
        .setRanges([matchupCell])
        .build();
      newRules.push(awayWin);

      // Margin column color (static) and conditional formatting
      marginCell = sheet.getRange(start,margins.getColumn(),end-start,1);
      marginCell.setBackground(dayColorsObj[contests[a].dayName]);
      const marginColorScaleRule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue(dayColorsObj[contests[a].dayName], SpreadsheetApp.InterpolationType.NUMBER, '1')
        .setGradientMaxpointWithValue(dayColorsFilledObj[contests[a].dayName], SpreadsheetApp.InterpolationType.NUMBER, '10')
        .setRanges([marginCell])
        .build();
      newRules.push(marginColorScaleRule);

      start = end;

    }
    end++;
  }
  sheet.getRange(startRow,margins.getColumn(),contests.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from({ length: 46 }, (_, index) => index), true).build());
  
  let allRules = rulesToKeep.concat(newRules);
  //clear all rules first and then add again
  
  sheet.setConditionalFormatRules(allRules);

  if (config.pickemsInclude) {
    // This function subcomponent runs when there are pick 'ems present and ties the response cell in Outcomes sheet to the response within the weekly sheet in question. It also prevents overwriting the values that may exist in the outcomes sheet, if present.
    let weeklySheetName = (weeklySheetPrefix + week);
    
    let sourceSheet = ss.getSheetByName(weeklySheetName);
    const targetSheet = ss.getSheetByName(`${LEAGUE}_OUTCOMES`);

    const sourceMatchupRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}`);
    const sourceMarginRange = ss.getRangeByName(`${LEAGUE}_PICKEM_OUTCOMES_${week}_MARGIN`);
    const targetMatchupRange = targetSheet.getRange(startRow,matchups.getColumn(),contests.length,1);
    const targetMarginRange = targetSheet.getRange(startRow,margins.getColumn(),contests.length,1);
    ss.setNamedRange(`${LEAGUE}_OUTCOMES_${week}`,targetMatchupRange);
    ss.setNamedRange(`${LEAGUE}_OUTCOMES_${week}_MARGIN`,targetMarginRange);

    let matchupCol = targetSheet.getRange(startRow,matchups.getColumn(),targetSheet.getMaxRows()-startRow+1,1);
    let dataMatchups = matchupCol.getValues().flat();
    let marginCol = targetSheet.getRange(startRow,margins.getColumn(),targetSheet.getMaxRows()-startRow+1,1);
    let dataMargins = marginCol.getValues().flat();
    let regexMatchups = new RegExp(/^[A-Z]{2,3}/);
    let regexMargins = new RegExp(/^[0-9]{1,2}/);
    let reWriteMatchups = [], reWriteMargins = [];
    // Data retention (if present)
    let map = [];
    for (let a = 0; a <= dataMatchups.length; a++) {
      if (regexMatchups.test(dataMatchups[a])) {
        reWriteMatchups.push(dataMatchups[a]);
        reWriteMargins.push(dataMargins[a]);
        map.push('');
      }
    }

    for (let a = 0; a < reWriteMatchups.length; a++) {
      if (teams.indexOf(reWriteMatchups[a])) {
        map[a] = Math.floor(teams.indexOf(reWriteMatchups[a])/2);
      }
    }
    matchupCol.clearContent();
    marginCol.clearContent();

    for (let e = map.length - 1; e >= 0; e--) {
      targetSheet.getRange(targetMatchupRange.getRow()+map[e],targetMatchupRange.getColumn()).setValue(reWriteMatchups[e]);
      targetSheet.getRange(targetMarginRange.getRow()+map[e],targetMarginRange.getColumn()).setValue(reWriteMargins[e]);
    }

    for (let a = 1; a <= sourceMatchupRange.getNumColumns(); a++) {
      if (!regexMatchups.test(dataMatchups[a-1])) {
        targetSheet.getRange(targetMatchupRange.getRow()+(a-1),targetMatchupRange.getColumn()).setFormula(
          '=\''+weeklySheetName+'\'!'+sourceSheet.getRange(sourceMatchupRange.getRow(),sourceMatchupRange.getColumn()+(a-1)).getA1Notation()
        );
      } else {
        Logger.log(`Found matching matchup outcome value of ${dataMatchups[a-1]} on outcomes sheet in row ${(a + 2)}; avoiding re-writing formula for this cell`);
      }
    }
    for (let a = 1; a <= sourceMarginRange.getNumColumns(); a++) {
      if (!regexMargins.test(dataMargins[a-1])) {
        targetSheet.getRange(targetMarginRange.getRow()+(a-1),targetMarginRange.getColumn()).setFormula(
          '=\''+weeklySheetName+'\'!'+sourceSheet.getRange(sourceMarginRange.getRow(),sourceMarginRange.getColumn()+(a-1)).getA1Notation()
        );
      } else {
        Logger.log(`Found matching outcome margin value of ${dataMargins[a-1]} on outcomes sheet in row ${(a + 2)}; avoiding re-writing formula for this cell`);
      }
    }
  }
}

/** 
 * TOTAL Sheet Creation / Adjustment
*/
function totSheet(ss,memberData) {
  ss = fetchSpreadsheet(ss);
  
  let docProps;
  if (!memberData) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;
  
  let sheetName = 'TOTAL';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);
  
  let rows = totalMembers+2;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }

  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  if ( weeks.length + 2 < maxCols ) {
    sheet.deleteColumns(weeks.length + 2,maxCols-(weeks.length + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('CORRECT');
  sheet.getRange(1,2).setValue('TOTAL');
  sheet.getRange(2,1).setValue('AVERAGES');

  for ( let a = 0; a < weeks.length; a++ ) {
    sheet.getRange(1,a+3).setValue(weeks[a]);
    sheet.setColumnWidth(a+3,30);
    sheet.getRange(2,a+3).setFormula('=iferror(arrayformula(countif(filter('+LEAGUE+'_PICKS_'+(weeks[a])+',NAMES_'+(weeks[a])+'=$A2)='+LEAGUE+'_PICKEM_OUTCOMES_'+(weeks[a])+',true)),)');
  }
  
  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,rows,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  sheet.getRange(rows,1,1,weeks.length+2).setBackground('#e6e6e6');
  
  sheet.getRange(2,2,totalMembers+1,weeks.length+1).setNumberFormat('#.#');

  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1); 

  // SET OVERALL NAMES Range
  let rangeOverallTotNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('TOT_OVERALL_NAMES',rangeOverallTotNames); 
  sheet.clearConditionalFormatRules(); 
  // OVERALL TOTAL GRADIENT RULE
  let rangeOverallTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL',rangeOverallTot);
  let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL"))') // Min value of all correct picks
    .setRanges([rangeOverallTot])
    .build();
  // OVERALL SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks.length+2));
  ss.setNamedRange('TOT_WEEKLY',range);
  let formatRuleOverallHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)>0,indirect(\"R[0]C[0]\",false)=max(indirect(\"R2C[0]:R'+maxRows+'C[0]\",false)))')
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, "15")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "10")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "5")
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallHigh);
  formatRules.push(formatRuleOverall);
  formatRules.push(formatRuleOverallTot);
  sheet.setConditionalFormatRules(formatRules);
  
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
  overallMainFormulas(weeks,sheet,totalMembers,'TOT',true);
  
  return sheet;  
}

// RNK Sheet Creation / Adjustment
function rnkSheet(ss,memberData) {
  ss = fetchSpreadsheet(ss);
  
  let docProps;
  if (!memberData) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;

  let sheetName = 'RNK';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.clear();
  sheet.setTabColor(generalTabColor);

  let rows = totalMembers + 1;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  if ( weeks.length + 2 < maxCols ) {
    sheet.deleteColumns(weeks.length + 2,maxCols-(weeks.length + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('RANKS');
  sheet.getRange(1,2).setValue('AVERAGE');

  for ( let a = 0; a < weeks.length; a++ ) {
    sheet.getRange(1,a+3).setValue(weeks[a]);
    sheet.setColumnWidth(a+3,48);
  }
    
  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,totalMembers+1,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  // SET OVERALL RANK NAMES Range
  let rangeOverallTotRnkNames = sheet.getRange('R2C1:R'+rows+'C1');
  ss.setNamedRange('TOT_OVERALL_RNK_NAMES',rangeOverallTotRnkNames);  
  sheet.clearConditionalFormatRules(); 
  // RANKS TOTAL GRADIENT RULE
  let rangeOverallRankTot = sheet.getRange('R2C2:R'+rows+'C2');
  ss.setNamedRange('TOT_OVERALL_RANK',rangeOverallRankTot);
  let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([rangeOverallRankTot])
    .build();
  // RANKS SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks.length+2));
  ss.setNamedRange('TOT_WEEKLY_RANK',range);
  let formatRuleOverallWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#00E1FF')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverall = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallWinner);
  formatRules.push(formatRuleOverall);
  formatRules.push(formatRuleOverallTot);
  sheet.setConditionalFormatRules(formatRules);
  
  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',false);
  overallMainFormulas(weeks,sheet,totalMembers,'RANK',false);
  
  return sheet;  
}

// PCT Sheet Creation / Adjustment
function pctSheet(ss,memberData) {
  ss = fetchSpreadsheet(ss);

  let docProps;
  if (!memberData) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;

  let sheetName = 'PCT';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);
  
  let rows = totalMembers+2; // 2 additional rows
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  if ( weeks.length + 2 < maxCols ) {
    sheet.deleteColumns(weeks.length + 2,maxCols-(weeks.length + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('PERCENTAGES');
  sheet.getRange(1,2).setValue('AVERAGE');
  sheet.getRange(rows,1).setValue('AVERAGES');
  
  for ( let a = 0; a < weeks.length; a++ ) {
    sheet.getRange(1,a+3).setValue(weeks[a]);
    sheet.setColumnWidth(a+3,48);
  }
  
  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,rows,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);
  
  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  sheet.getRange(rows,1,1,weeks.length+2).setBackground('#e6e6e6'); 

  sheet.getRange(2,2,totalMembers+1,1).setNumberFormat("##.#%");  
  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1);

  // SET OVERALL PCT NAMES Range
  let rangeOverallTotPctNames = sheet.getRange('R2C1:R'+(rows-1)+'C1');
  ss.setNamedRange('TOT_OVERALL_PCT_NAMES',rangeOverallTotPctNames);
  sheet.clearConditionalFormatRules();
  // PCT TOTAL GRADIENT RULE
  let rangeOverallTotPct = sheet.getRange('R2C2:R'+(rows-1)+'C2');
  ss.setNamedRange('TOT_OVERALL_PCT',rangeOverallTotPct);
  rangeOverallTotPct = sheet.getRange('R2C2:R'+rows+'C2');
  let formatRuleOverallPctTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("TOT_OVERALL_PCT"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("TOT_OVERALL_PCT"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("TOT_OVERALL_PCT"))') // Min value of all correct picks  
    .setRanges([rangeOverallTotPct])
    .build();  
  // PCT SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+(rows-1)+'C'+(weeks.length+2));
  ss.setNamedRange('TOT_WEEKLY_PCT',range);
  range = sheet.getRange('R2C3:R'+rows+'C'+(weeks.length+2)); 
  let formatRuleOverallPctHigh = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(indirect(\"R[0]C[0]\",false)>0,indirect(\"R[0]C[0]\",false)=max(indirect(\"R2C[0]:R'+maxRows+'C[0]\",false)))')
    .setBackground('#75F0A1')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOverallPct = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleOverallPctHigh);
  formatRules.push(formatRuleOverallPct);
  formatRules.push(formatRuleOverallPctTot);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',true);
  overallMainFormulas(weeks,sheet,totalMembers,'PCT',true);

  return sheet;  
}

// MNF Sheet Creation / Adjustment
function mnfSheet(ss,memberData) {
  ss = fetchSpreadsheet(ss);

  let docProps;
  if (!memberData) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;

  let sheetName = 'MNF';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  }

  sheet.clear();
  sheet.setTabColor(generalTabColor);
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));

  Logger.log('Checking for Monday games, if any');
  let data = ss.getRangeByName(LEAGUE).getValues();
  let text = '0';
  let result = text.repeat(weeks.length);
  let mondayNightGames = Array.from(result);
  for (let a = 0; a < data.length; a++) {
    if ( data[a][2] == 1 && data[a][3] >= 17) {
      mondayNightGames[(data[a][0]-1)]++;
    }
  }
  let rows = totalMembers + 2; // AustinOrphan's suggestion!
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( weeks.length + 2 < maxCols ) {
    sheet.deleteColumns(weeks.length + 2,maxCols-(weeks.length + 2));
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue('CORRECT');
  sheet.getRange(1,2).setValue('TOTAL');
  sheet.getRange(rows,1).setValue('AVERAGES');

  let range = sheet.getRange(1,1,rows,maxCols);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,rows,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  sheet.setColumnWidth(2,70);

  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  sheet.getRange(rows,1,1,weeks.length+2).setBackground('#e6e6e6'); 
  
  let headers = [];
  for ( let a = 0; a < weeks.length; a++ ) {
    if (mondayNightGames[a] == 2) {
      range = sheet.getRange(1,a+3);
      range.setNote('Two MNF Games')
        .setFontWeight('bold')
        .setBackground('#555555');
    } else if (mondayNightGames[a] == 3) {
      range = sheet.getRange(1,a+3);
      range.setNote('Three MNF Games')
        .setFontWeight('bold')
        .setBackground('#999999');
    } else if (mondayNightGames[a] == 4) {
      range = sheet.getRange(1,a+3);
      range.setNote('Four MNF Games')
        .setFontWeight('bold')
        .setBackground('#CCCCCC');
    } else if (mondayNightGames[a] >= 4) {
      range = sheet.getRange(1,a+3);
      range.setNote(mondayNightGames[a] + ' MNF Games')
        .setFontWeight('bold')
        .setFontColor('black')
        .setBackground('#FFFFFF');
    }
    sheet.setColumnWidth(a+3,30);
    headers.push(weeks[a]);
  }
  sheet.getRange(1,3,1,weeks.length).setValues([headers]);

  sheet.setFrozenColumns(2);
  sheet.setFrozenRows(1); 

  sheet.clearConditionalFormatRules(); 

  // SET MNF NAMES Range
  let rangeMnfNames = sheet.getRange('R2C1:R'+(rows-1)+'C1');
  ss.setNamedRange('MNF_NAMES',rangeMnfNames); 
  // MNF TOTAL GRADIENT RULE
  let rangeMnfTot = sheet.getRange('R2C2:R'+(rows-1)+'C2');
  ss.setNamedRange('MNF',rangeMnfTot);
  let formatRuleMnfTot = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#C9FFDF", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect("MNF"))') // Max value of all correct picks
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect("MNF"))') // Generates Median Value
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect("MNF"))') // Min value of all correct picks
    .setRanges([rangeMnfTot])
    .build();
  // MNF AVERAGES GRADIENT RULE
  let rangeMnfAvg = sheet.getRange('R'+rows+'C2:R'+rows+'C'+(weeks.length+2));
  let formatRuleMnfAvg = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#C9FFDF", SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "0.5")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setRanges([rangeMnfAvg])
    .build();
  // MNF SHEET GRADIENT RULE
  range = sheet.getRange('R2C3:R'+(rows-1)+'C'+(weeks.length+2));
  ss.setNamedRange('MNF_WEEKLY',range);
  let formatRuleTwoCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(2)
    .setBackground('#9CFFC4')
    .setFontColor('#9CFFC4')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleOneCorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#C9FFDF')
    .setFontColor('#C9FFDF')
    .setBold(true)
    .setRanges([range])
    .build();
  let formatRuleIncorrect = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=or(and(not(isblank(indirect(\"R[0]C[0]\",false))),indirect(\"R[0]C[0]\",false)=0),and(isblank(indirect(\"R[0]C[0]\",false)),indirect(\"WEEK\")>=indirect(\"R1C[0]\",false)))')
    .setBackground('#FFC4CA')
    .setFontColor('#FFC4CA')
    .setBold(true)
    .setRanges([range])
    .build();    
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(formatRuleTwoCorrect);
  formatRules.push(formatRuleOneCorrect);
  formatRules.push(formatRuleIncorrect);
  formatRules.push(formatRuleMnfTot);
  formatRules.push(formatRuleMnfAvg);
  sheet.setConditionalFormatRules(formatRules);

  overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',false);
  overallMainFormulas(weeks,sheet,totalMembers,'MNF',true);

  return sheet;  
}


/**
 * Sheet creation tool for the survivor and eliminator sheets
 * 
 */
function survElimSheet(ss,config,memberData,sheetType) {
  ss = ss || fetchSpreadsheet(ss);
  let docProps;
  if (!config || !memberData) docProps = PropertiesService.getDocumentProperties();

  config = config || JSON.parse(docProps.getProperty('configuration')) || {};
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  
  sheetType = sheetType || 'survivor'; // Default to survivor
  const sheetName = sheetType.toUpperCase();

  let sheet = ss.getSheetByName(sheetName);
  let fresh = false;
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    fresh = true;
  }

  sheet.setTabColor(survElimTabColors[sheetType]);

  const totalMembers = memberData.memberOrder.length;
  const members = memberData.memberOrder.map(id => [memberData.members[id]?.name]);

  let maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();

  let previousDataRange, previousData;
  if (!fresh){
    previousDataRange = sheet.getRange(2,3,maxRows-2,WEEKS - WEEKS_TO_EXCLUDE.length);
    previousData = previousDataRange.getValues();
    const text = `💾 Gathered previous data for ${sheetName} sheet, recreating sheet now`;
    Logger.log(text);
    ss.toast(text,`${sheetName} BACKED UP`);
  }
  sheet.clear();

  let rows = totalMembers + 2;
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let cols = WEEKS - WEEKS_TO_EXCLUDE.length + 2;
  if (cols < maxCols) {
    sheet.deleteColumns(cols + 1,maxCols-cols);
  } else if (cols > maxCols) {
    sheet.insertColumnsAfter(maxCols,cols-maxCols);
  }
  maxCols = sheet.getMaxColumns();
  
  sheet.getRange(1,1).setValue('PLAYER');
  let livesCol = 2;
  sheet.getRange(1,livesCol).setValue('LIVES');
  sheet.setColumnWidth(livesCol,75);
  let revivesCol = 3;
  sheet.getRange(1,revivesCol).setValue('REVIVES');
  sheet.setColumnWidth(revivesCol,50);
  let eliminatedCol = 4;
  sheet.getRange(1,eliminatedCol).setValue('ELIMINATED');
  sheet.setColumnWidth(eliminatedCol,100);
  
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  
  for (let a = 0; a < weeks.length; a++ ) {
    sheet.getRange(1,a+5).setValue(weeks[a]);
    sheet.setColumnWidth(a+5,30);
  }

  let range = sheet.getRange(1,1,rows,weeks.length+2);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(members);
  sheet.getRange(totalMembers+2,1).setValue('REMAINING');
  sheet.getRange(1,1,totalMembers+2,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,120);
  
  range = sheet.getRange(1,1,1,weeks.length+4);
  range.setBackground('black');
  range.setFontColor('white');
  range = sheet.getRange(totalMembers+2,1,1,weeks.length+4);
  range.setBackground('#e6e6e6');
  
  sheet.setFrozenColumns(4);
  sheet.setFrozenRows(1);
  
  ss.setNamedRange(`${sheetName}_NAMES`,sheet.getRange(2,1,totalMembers,1));
  ss.setNamedRange(`${sheetName}_LIVES`,sheet.getRange(2,2,totalMembers,1))
  ss.setNamedRange(`${sheetName}_REVIVES`,sheet.getRange(2,3,totalMembers,1))
  ss.setNamedRange(`${sheetName}_ELIMINATED`,sheet.getRange(2,4,totalMembers,1))
  ss.setNamedRange(`${sheetName}_PICKS`,sheet.getRange(2,5,totalMembers,weeks.length));

  if (config[`${sheetType}Lives`] == 1) sheet.hideColumns(livesCol);
  if (!config[`${sheetType}Revives`]) sheet.hideColumns(revivesCol);

  if (!fresh) {
    previousDataRange.setValues(previousData);
    const text = `🔄 Previous values restored for ${sheetName} sheet if they were present`;
    Logger.log(text);
    ss.toast(text,`${sheetName} RESTORED`);
  }
  return sheet;
}

// WINNERS Sheet Creation / Adjustment
function winnersSheet(ss,year) {
  ss = fetchSpreadsheet(ss);

  let sheetName = 'WINNERS';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  sheet.clear();
  sheet.setTabColor(winnersTabColor);
  
  let checkboxRange = sheet.getRange(2,3,WEEKS+3,1);
  let checkboxes = checkboxRange.getValues();
  
  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));

  let rows = weeks.length + 5;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( 3 < maxCols ) {
    sheet.deleteColumns(3,maxCols-3);
  }
  maxCols = sheet.getMaxColumns();
  sheet.getRange(1,1).setValue(year);
  sheet.getRange(1,2).setValue('WINNER');
  sheet.getRange(1,3).setValue('PAID');

  let range = sheet.getRange(1,1,rows,maxCols);
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,2,rows-1,1).setHorizontalAlignment('left');
  sheet.setColumnWidth(1,80);
  sheet.setColumnWidth(2,150);
  sheet.setColumnWidth(3,40);

  range = sheet.getRange(2,3,rows-1,1);
  range.insertCheckboxes();
  range.setHorizontalAlignment('center');
  range = sheet.getRange(1,1,rows,2);
  range.setHorizontalAlignment('left');
  let a = 0;
  for (a; a <= weeks.length; a++) {
    sheet.getRange(a+2,1,1,1).setValue(weeks[a]);
  }
  sheet.getRange(a+1,1,4,1).setValues([['SURVIVOR'],['ELIMINATOR'],['MNF'],['OVERALL']]);

  range = sheet.getRange(1,1,1,maxCols);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenRows(1); 

  range = sheet.getRange('R2C2:R'+(weeks.length+1)+'C2');
  ss.setNamedRange('WEEKLY_WINNERS',range);

  sheet.clearConditionalFormatRules(); 
  // OVERALL SHEET GRADIENT RULE
  let fivePlusWins = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=countif($2:B$'+(weeks.length+1)+',B2)>=5')
  .setBackground('#2CFF75')
  .setRanges([range])
  .build();
  let fourPlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks.length+1)+',B2)=4')
    .setBackground('#72FFA3')
    .setRanges([range])
    .build();
  let threePlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks.length+1)+',B2)=3')
    .setBackground('#BBFFD3')
    .setRanges([range])
    .build();
  let twoPlusWins = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=countif(B$2:B$'+(weeks.length+1)+',B2)=2')
    .setBackground('#D3FFE2')
    .setRanges([range])
    .build();
  let formatRules = sheet.getConditionalFormatRules();
  formatRules.push(fivePlusWins);
  formatRules.push(fourPlusWins);
  formatRules.push(threePlusWins);
  formatRules.push(twoPlusWins);
  sheet.setConditionalFormatRules(formatRules);
  
  // Rewrites the checkboxes if they previously had any checked.
  let col = checkboxRange.getColumn();
  for (let a = 0; (a < checkboxes.length || a < (weeks.length + 4)); a++) {
    if (checkboxes[a][0]) {
      sheet.getRange(a+1,col).check();
    }
  }
  let winRange;
  let nameRange;

  for ( let b = 1; b <= weeks.length; b++ ) {
    winRange = 'WIN_' + (b);
    nameRange = 'NAMES_' + (b);
    sheet.getRange(b+1,2,1,1).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }

  return sheet;

}

// SUMMARY Sheet Creation / Adjustment
function summarySheet(ss,memberData,config) {
  ss = fetchSpreadsheet(ss);
  let restoreNotes = false;
  let notesRange, notes, sheetName = 'SUMMARY';
  
  let docProps;
  if (!memberData || !config) docProps = PropertiesService.getDocumentProperties();
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  const memberNames = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  const totalMembers = memberNames.length;
  
  config = config || JSON.parse(docProps.getProperty('configuration')) || {};

  let sheet = ss.getSheetByName(sheetName);
  if (sheet == null) {
    ss.insertSheet(sheetName);
    sheet = ss.getSheetByName(sheetName);
  } else {
    restoreNotes = true;
    notesRange = sheet.getRange(2,sheet.getRange(1,1,sheet.getMaxRows()-1,sheet.getMaxColumns()).getValues().flat().indexOf('NOTES')+1,sheet.getMaxRows()-1,1);
    notes = notesRange.getValues();
  }
  sheet.clear();
  sheet.setTabColor(winnersTabColor);

  let headers = ['PLAYER'];
  let headersWidth = [120];
  let mnfCol;
  if (config.pickemsInclude) {
    headers = headers.concat(['TOTAL CORRECT','TOTAL RANK','AVG % CORRECT','AVG % CORRECT RANK','WEEKLY WINS']);
    headersWidth = headersWidth.concat([90,90,90,90,90]);
    if (!config.mnfExclude) {
      headers = headers.concat(['MNF CORRECT','MNF RANK']);
      headersWidth = headersWidth.concat([90,90]);
      mnfCol = headers.indexOf('MNF CORRECT') + 1;
    }
  }

  let survivorCol,eliminatorCol
  if (config.survivorInclude) {
    headers.push('SURVIVOR LIVES');
    headers.push('SURVIVOR (WEEK OUT)');
    headersWidth.push(65);
    headersWidth.push(90);
    survivorCol = headers.indexOf('SURVIVOR (WEEK OUT)')+1;
  }
  if (config.eliminatorInclude) {
    headers.push('ELIMINATOR LIVES');
    headers.push('ELIMINATOR (WEEK OUT)');
    headersWidth.push(65);
    headersWidth.push(90);
    eliminatorCol = headers.indexOf('ELIMINATOR (WEEK OUT)')+1;
  }
  headers.push('NOTES');
  headersWidth.push(160);
  
  let totalCol = headers.indexOf('TOTAL CORRECT') + 1;
  let weeklyPercentCol = headers.indexOf('AVG % CORRECT') + 1;
  let weeklyRankAvgCol = headers.indexOf('AVG % CORRECT RANK') + 1;
  let weeklyWinsCol = headers.indexOf('WEEKLY WINS') + 1;
  let notesCol = headers.indexOf('NOTES') + 1;

  let len = headers.length;
  
  let rows = totalMembers + 1;
  let maxRows = sheet.getMaxRows();
  if (rows < maxRows) {
    sheet.deleteRows(rows,maxRows-rows);
  } else if (rows > maxRows){
    sheet.insertRows(maxRows,rows-maxRows);
  }
  maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  if ( len < maxCols ) {
    sheet.deleteColumns(len,maxCols-len);
  } else if ( len > maxCols ) {
    sheet.insertColumnsAfter(maxCols, len - maxCols);
  }
  maxCols = sheet.getMaxColumns();
  
  sheet.getRange(1,1,1,len).setValues([headers]);
  if(restoreNotes) {
    sheet.getRange(2,notesCol,notes.length,1).setValues(notes);
  }
  
  for ( let a = 0; a < len; a++ ) {
    sheet.setColumnWidth(a+1,headersWidth[a]);
  }
  sheet.setRowHeight(1,40);
  let range = sheet.getRange(1,1,1,maxCols);
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  range = sheet.getRange(1,1,rows,len);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  range.setFontFamily("Montserrat");
  range.setFontSize(10);
  sheet.getRange(2,1,totalMembers,1).setValues(memberNames); 
  sheet.getRange(1,1,totalMembers+1,1).setHorizontalAlignment('left');
  
  range = sheet.getRange(1,1,1,len);
  range.setBackground('black');
  range.setFontColor('white');
  
  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(1);
  
  sheet.clearConditionalFormatRules(); 
  let formatRules = sheet.getConditionalFormatRules();
  if (config.pickemsInclude) {
    // SUMMARY TOTAL GRADIENT RULE
    let rangeSummaryTot = sheet.getRange('R2C'+totalCol+':R'+rows+'C'+totalCol);
    let formatRuleOverallTot = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint('#75F0A1')
      .setGradientMinpoint('#FFFFFF')
      .setRanges([rangeSummaryTot])
      .build();
    formatRules.push(formatRuleOverallTot);
    // MNF TOTAL GRADIENT RULES
    let rangeMNFTot, rangeMNFRank, formatRuleMNFRank;
    if (config.mnfInclude) {
      rangeMNFTot = sheet.getRange('R2C'+mnfCol+':R'+rows+'C'+mnfCol);
      //ss.setNamedRange('TOT_MNF',range);
      let formatRuleMNFTot = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint('#75F0A1')
        .setGradientMinpoint('#FFFFFF')
        .setRanges([rangeMNFTot])
        .build();
      formatRules.push(formatRuleMNFTot);    
      // RANK MNF GRADIENT RULE
      rangeMNFRank = sheet.getRange('R2C'+(mnfCol+1)+':R'+rows+'C'+(mnfCol+1));
      ss.setNamedRange('TOT_MNF_RANK',rangeMNFRank);
      formatRuleMNFRank = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
        .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
        .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
        .setRanges([rangeMNFRank])
        .build();
      formatRules.push(formatRuleMNFRank);
    }
    // RANK OVERALL RULE
    let rangeOverallRank = sheet.getRange('R2C'+(totalCol+1)+':R'+rows+'C'+(totalCol+1));
    ss.setNamedRange('TOT_OVERALL_RANK',rangeOverallRank);
    let formatRuleRank = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))')
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=counta(indirect("MEMBERS"))/2')
      .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
      .setRanges([rangeOverallRank])
      .build();
    formatRules.push(formatRuleRank);
    // WEEKLY WINS GRADIENT/SINGLE COLOR RULES
    range = sheet.getRange('R2C'+weeklyWinsCol+':R'+rows+'C'+weeklyWinsCol);
    ss.setNamedRange('WEEKLY_WINS',range); 
    let formatRuleWeeklyWinsEmpty = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#FFFFFF')
      .setFontColor('#FFFFFF')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleWeeklyWinsEmpty);
    let formatRuleWeeklyWins = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint('#ffee00')
      .setGradientMinpoint('#FFFFFF')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleWeeklyWins);   
    // OVERALL AND WEEKLY CORRECT % AVG
    range = sheet.getRange('R2C'+weeklyPercentCol+':R'+rows+'C'+weeklyPercentCol);
    range.setNumberFormat('##.#%');
    let formatRuleCorrectAvg = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, ".70")
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".60")
      .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, ".50")
      .setRanges([range])
      .build();
    formatRules.push(formatRuleCorrectAvg);
    // WEEKLY RANK AVG
    range = sheet.getRange('R2C'+weeklyRankAvgCol+':R'+rows+'C'+weeklyRankAvgCol);
    range.setNumberFormat('#.#');
    let formatRuleCorrectRank = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, "5")
      .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, "10")
      .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, "15")
      .setRanges([range])
      .build();
    formatRules.push(formatRuleCorrectRank);
  }
  if (config.survivorInclude) {
  // SURVIVOR "IN"
    range = sheet.getRange('R2C'+survivorCol+':R'+(totalMembers+1)+'C'+survivorCol);
    let formatRuleIn = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('IN')
      .setBackground('#C9FFDF')
      .setRanges([range])
      .build();
    // SURVIVOR "OUT"
    let formatRuleOut = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OUT')
      .setBackground('#F2BDC2')
      .setRanges([range])
      .build();    
    formatRules.push(formatRuleIn);
    formatRules.push(formatRuleOut);
  }
  if (config.eliminatorInclude) {
  // ELIMINATOR "IN"
    range = sheet.getRange('R2C'+eliminatorCol+':R'+(totalMembers+1)+'C'+eliminatorCol);
    let formatRuleIn = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('IN')
      .setBackground('#C9FFDF')
      .setRanges([range])
      .build();
    // ELIMINATOR "OUT"
    let formatRuleOut = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('OUT')
      .setBackground('#F2BDC2')
      .setRanges([range])
      .build();    
    formatRules.push(formatRuleIn);
    formatRules.push(formatRuleOut);
  }  
  sheet.setConditionalFormatRules(formatRules);
  // Creates all formulas for SUMMARY Sheet
  summarySheetFormulas(headers, sheet,totalMembers);

  return sheet;  
}

// UPDATES SUMMARY SHEET FORMULAS
function summarySheetFormulas(headers,sheet,totalMembers) {
  let arr = [...headers] || ['PLAYER','TOTAL CORRECT','TOTAL RANK','MNF CORRECT','MNF RANK','AVG % CORRECT','AVG % CORRECT RANK','WEEKLY WINS','SURVIVOR LIVES','SURVIVOR (WEEK OUT)','ELIMINATOR LIVES','ELIMINATOR (WEEK OUT)','NOTES'];
  
  if (!sheet) {
    sheet = fetchSpreadsheet().getSheetByName('SUMMARY');  
  }
  headers.unshift('COL INDEX ADJUST');

  for (let a = 0; a < arr.length; a++) {
    for (let b = 0; b < totalMembers; b++) {
      if (headers[a] == 'TOTAL CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_NAMES,TOT_OVERALL},2,false))');
      } else if (headers[a] == 'TOTAL RANK' || headers[a] == 'AVG % CORRECT RANK' || headers[a] == 'MNF RANK') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(rank(R[0]C[-1],R2C[-1]:R'+ (totalMembers+1) + 'C[-1]))');
        ss.setNamedRange('TOT_OVERALL_RANK',sheet.getRange(2,headers.indexOf('TOTAL RANK'),totalMembers,1));
      } else if (headers[a] == 'MNF CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{MNF_NAMES,MNF},2,false))');
        ss.setNamedRange('TOT_MNF_RANK',sheet.getRange(2,headers.indexOf('MNF RANK'),totalMembers,1));
      } else if (headers[a] == 'AVG % CORRECT') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(vlookup(R[0]C1,{TOT_OVERALL_PCT_NAMES,TOT_OVERALL_PCT},2,false))');
      } else if (headers[a] == 'WEEKLY WINS') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(countif(WEEKLY_WINNERS,R[0]C1))');
        ss.setNamedRange('WEEKLY_WINS',sheet.getRange(2,headers.indexOf('WEEKLY WINS'),totalMembers,1));
      } else if (headers[a] == 'SURVIVOR (WEEK OUT)') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(arrayformula(if(isblank(vlookup(R[0]C1,{SURVIVOR_NAMES,SURVIVOR_ELIMINATED},2,false)),"IN","OUT ("\&vlookup(R[0]C1,{SURVIVOR_NAMES,SURVIVOR_ELIMINATED},2,false)\&")")))');
      } else if (headers[a] == 'ELIMINATOR (WEEK OUT)') {
        sheet.getRange(b+2,a).setFormulaR1C1('=iferror(arrayformula(if(isblank(vlookup(R[0]C1,{ELIMINATOR_NAMES,ELIMINATOR_ELIMINATED},2,false)),"IN","OUT ("\&vlookup(R[0]C1,{ELIMINATOR_NAMES,ELIMINATOR_ELIMINATED},2,false)\&")")))');
      }
    }
  }
  Logger.log('Updated formulas and ranges for summary sheet');
}

// TOT / RANK / PCT / MNF Combination formula for sum/average per player row
function overallPrimaryFormulas(sheet,totalMembers,maxCols,action,avgRow) {
  if (action == 'average') {
    sheet.getRange(2,2,totalMembers,1).setFormulaR1C1('=iferror(if(counta(R[0]C3:R[0]C'+maxCols+')=0,,average(R[0]C3:R[0]C'+maxCols+')))')
      .setNumberFormat("#0.0");
  } else if (action == 'sum') {
    sheet.getRange(2,2,totalMembers,1).setFormulaR1C1('=iferror(if(counta(R[0]C3:R[0]C'+maxCols+')=0,,sum(R[0]C3:R[0]C'+maxCols+')))')
      .setNumberFormat("##");
  }
  if (sheet.getSheetName() == 'PCT') {
    sheet.getRange(2,2,totalMembers,1).setNumberFormat("##.#%");
  }
  if (avgRow) {
    if (sheet.getSheetName() == 'PCT'){  
      sheet.getRange(sheet.getMaxRows(),2).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
        .setNumberFormat('##.#%');
    } else {
      sheet.getRange(sheet.getMaxRows(),2).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
        .setNumberFormat("#0.0");
    }
  }
}

// TOT / RNK / PCT / MNF Combination formula for each column (week)
function overallMainFormulas(weeks,sheet,totalMembers,str,avgRow) {
  let b;
  weeks = weeks || Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));
  for (let a = 0; a < weeks.length; a++ ) {
    b = 1;
    for (b ; b <= totalMembers; b++) {
      if (str == 'TOT') {
        sheet.getRange(b+1,a+3).setFormula('=iferror(if(or(iserror(vlookup($A'+(b+1)+',NAMES_'+weeks[a]+',1,false)),counta(filter('+LEAGUE+'_PICKS_'+weeks[a]+',NAMES_'+weeks[a]+'=$A'+(b+1)+'))=0),,arrayformula(countifs(filter('+LEAGUE+'_PICKS_'+weeks[a]+',NAMES_'+weeks[a]+'=$A'+(b+1)+')='+LEAGUE+'_PICKEM_OUTCOMES_'+weeks[a]+',true,filter('+LEAGUE+'_PICKS_'+weeks[a]+',NAMES_'+weeks[a]+'=$A'+(b+1)+'),\"<>\"))),)');
      } else {
        sheet.getRange(b+1,a+3).setFormula('=iferror(arrayformula(vlookup(R[0]C1,{NAMES_'+weeks[a]+','+str+'_'+weeks[a]+'},2,false)))');
      }
      if (sheet.getSheetName() == 'PCT') {
        sheet.getRange(b+1,a+3).setNumberFormat("##.#%");
      } else {
        sheet.getRange(b+1,a+3).setNumberFormat("#0");
      }
    }
  }
  if (avgRow) {
    if (sheet.getSheetName() == 'MNF') {
      // Instance of MNF sheet, where sheet needs to have data for quantity of MNF games
      Logger.log('Checking for Monday games, if any');
      let data = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(LEAGUE).getValues();
      let text = '0';
      let result = text.repeat(weeks.length);
      let mondayNightGames = Array.from(result);
      for (let a = 0; a < data.length; a++) {
        if ( data[a][2] == 1 && data[a][3] >= 17) {
          mondayNightGames[(data[a][0]-1)]++;
        }
      }
      for (let a = 0; a < weeks.length; a++){
        let rows = sheet.getMaxRows();
        if (mondayNightGames[a] > 1) {
          sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0])/'+mondayNightGames[a]+',))')
            .setNumberFormat("##%");
        } else {
          sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))')
            .setNumberFormat("##%");
        }
      }
    } else {
      for (let a = 0; a < weeks.length; a++){
        let rows = sheet.getMaxRows();
        sheet.getRange(rows,a+3).setFormulaR1C1('=iferror(if(counta(R2C[0]:R'+(totalMembers+1)+'C[0])>=3,average(R2C[0]:R'+(totalMembers+1)+'C[0]),))');
      }
    }
  }
}

// WEEKLY WINNERS Combination formula update
function winnersFormulas(weeks,sheet) {
  for (let a = 0; a < weeks.length; a++ ) {
    let winRange = `WIN_${weeks[a]}`;
    let nameRange = `NAMES_${weeks[a]}`;
    sheet.getRange(a+1,2).setFormulaR1C1('=iferror(join(", ",sort(filter('+nameRange+','+winRange+'=1),1,true)))');
  }
}

// REFRESH FORMULAS FOR TOT / RNK / PCT / MNF
function allFormulasUpdate(ss){
  ss = fetchSpreadsheet(ss);
  const docProps = PropertiesService.getDocumentProperties();
  const config = JSON.parse(docProps.getProperty('configuration')) || {};
  const memberData = JSON.parse(docProps.getProperty('members')) || {};

  const totalMembers = memberData.memberOrder.length();
  let sheet, maxCols;

  const weeks = Array.from({ length: WEEKS }, (_, index) => index + 1).filter(week => !WEEKS_TO_EXCLUDE.includes(week));

  if (config.pickemsInclude) {
    sheet = ss.getSheetByName('TOTAL');
    maxCols = sheet.getMaxColumns();
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
    overallMainFormulas(weeks,sheet,totalMembers,'TOT',true);

    sheet = ss.getSheetByName('RNK');
    maxCols = sheet.getMaxColumns();
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',false);
    overallMainFormulas(weeks,sheet,totalMembers,'RNK',false);
  
    sheet = ss.getSheetByName('PCT');
    maxCols = sheet.getMaxColumns();
    overallPrimaryFormulas(sheet,totalMembers,maxCols,'average',true);
    overallMainFormulas(weeks,sheet,totalMembers,'PCT',true);
    
    if (!config.mnfExclude) {
      sheet = ss.getSheetByName('MNF');
      maxCols = sheet.getMaxColumns();
      overallPrimaryFormulas(sheet,totalMembers,maxCols,'sum',true);
      overallMainFormulas(weeks,sheet,totalMembers,'MNF',true);
    }

    sheet = ss.getSheetByName('WINNERS');
    winnersFormulas(weeks,sheet);
  }
}

// ============================================================================================================================================
// WEEKLY SHEETS
// ============================================================================================================================================

// WEEKLY Sheet Function - creates a sheet with provided week, members [array], and if data should be restored
function weeklySheet(ss,week,config,forms,memberData,displayEmpty) {
  ss = ss || fetchSpreadsheet(ss);
  week = week || fetchWeek();
  let docProps = (!config || !forms || !memberData) ? PropertiesService.getDocumentProperties() : null;
  forms = forms || JSON.parse(docProps.getProperty('forms')) || {};

  if (!forms[week].gamePlan.pickemsInclude) {
    Logger.log(`⭕ Pick 'Ems not included in week ${week} form response, no weekly sheet needed`);
    ss.toast(`⭕ Pick 'Ems not included in week ${week} form response, no weekly sheet needed`);
    return null;
  }
  config = config || JSON.parse(docProps.getProperty('configuration')) || {};
  memberData = memberData || JSON.parse(docProps.getProperty('members')) || {};
  
  // Insert Members
  let members = [];
  if (!displayEmpty) { //|| forms[week].respondents == totalMembers) {
    members = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
  } else {
    members = memberData.memberOrder.map(id => [memberData.members[id]?.name]);
    members = [];

    // Re-sorts based on memberOrder, then applies conversion to the name
    const sortedRespondents = forms[week].respondents.sort((a, b) => {
      const indexA = memberData.memberOrder.indexOf(a);
      const indexB = memberData.memberOrder.indexOf(b);
      const resolvedIndexA = indexA > -1 ? indexA : Infinity;
      const resolvedIndexB = indexB > -1 ? indexB : Infinity;
      return resolvedIndexA - resolvedIndexB;
    });
    members = sortedRespondents.map((id) => [memberData.members[id]?.name]);
  }

  let totalMembers = members.length;
  
  if (totalMembers <= 0) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(`⚠️ MEMBER ISSUE`, `There was an issue fetching the members to create the weekly sheet, make sure you've used the "Member Management" panel or waited for first submissions of the form before creating this sheet`,ui.ButtonSet.OK);
    Logger.log('⚠️ Error fetching members to create weekly sheet');
    return null;
  }

  let sheet, sheetName = weeklySheetPrefix + week;
  const contests = forms[week].gamePlan.games;
  const isAts = forms[week].gamePlan.pickemsAts;
  
  let diffCount = (totalMembers - 1) >= 5 ? 5 : (totalMembers - 1); // Number of results to display for most similar weekly picks (defaults to 5, or 1 fewer than the total member count, whichever is larger)

  const matchRow = 1; // Row for all matchups
  const dayRow = matchRow + 1; // Row for denoting day of the week
  const entryRowStart = dayRow + 1; // Row of first user input on weekly sheet
  const entryRowEnd = (entryRowStart - 1) + totalMembers; // Includes any header rows (entryRowStart-1) and adds two additional for final row of home/away splits and then bonus values
  const summaryRow = entryRowEnd + 1; // Row for group averages (away/home) and other calculated values
  const spreadRow = summaryRow + 1; // Recorded spreads (hidden if not ATS)
  const outcomeRow = summaryRow + 2; // Row for matchup outcomes
  const outcomeMarginRow = summaryRow + 3; // Row for margins
  const spreadOutcomeRow = summaryRow + 4; // Row for determining which team was the corret pick when including the spread
  const bonusRow = summaryRow + 5; // Row for adding bonus drop-downs
  const rows = bonusRow; // Declare row variable, unnecessary, but easier to work with  
  const pointsCol = 2;

  let columns, fresh = false;
  
  // Checks for sheet presence and creates if necessary
  sheet = ss.getSheetByName(sheetName);  
  if (sheet == null) {
    dataRestore = false;
    ss.insertSheet(sheetName,ss.getNumSheets()+1);
    sheet = ss.getSheetByName(sheetName);
    fresh = true;
  }

  // Adds tab colors
  weeklySheetTabColors(ss,sheet); 

  let maxRows = sheet.getMaxRows();
  let maxCols = sheet.getMaxColumns();
  
  // DATA GATHERING IF DATA RESTORE ACTIVE
  let regex, commentCol, tiebreakerCol = -1, firstInput, finalInput, previousRows, previousNames, previousData, previousOutcomes, previousComment, previousTiebreaker, previousTiebreakers, previousBonus = null, previousNamesRange, previousDataRange, previousOutcomesRange, previousCommentRange, previousTiebreakersRange, matchupStartCol, matchupEndCol;
  
  // CLEAR AND PREP SHEET
  sheet.clear();
  sheet.clearNotes();
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).clearDataValidations();
  adjustRows(sheet,rows);
  
  sheet.getRange(entryRowStart,1,totalMembers,1).setValues(members);

  // Setting header values
  let headers = ['WEEK ' + week,'POINTS','WEEKLY\nRANK','PERCENT\nCORRECT'];
  let bottomHeaders = ['PREFERRED','AWAY','HOME'];
  sheet.getRange(summaryRow,1,1,3).setValues([bottomHeaders]);
  sheet.getRange(spreadRow,1).setValue('SPREAD VALUE');
  sheet.getRange(outcomeRow,1).setValue('WINNER');
  sheet.getRange(outcomeMarginRow,1).setValue('MARGIN OF VICTORY');
  sheet.getRange(spreadOutcomeRow,1).setValue('WINNER AGAINST THE SPREAD');
  sheet.getRange(bonusRow,1).setValue('BONUS');
  let widths = [130,75,75,75];

  // Setting headers for the week's matchups with format of 'AWAY' + '@' + 'HOME', then creating a data validation cell below each
  let firstMatchCol = headers.length + 1;
  let mnfCol, mnfStartCol, mnfEndCol, tnfStartCol, tnfEndCol, winCol, days = [], spreads = [], dayRowColors = [], bonuses = [], formatRules = [];
  let mnf = false, tnf = false; // Preliminary establishing if there are Monday or Thursday games (false by default, fixed to true when looped through matchups)
  let rule, matches = contests.length;
  for ( let a in contests ) {
    let day = contests[a].dayName;
    let evening = contests[a].hour >= 17 ? true : false;
    let away = contests[a].awayTeam;
    let home = contests[a].homeTeam;
    // Establish start/stop of MNF games to record the tally
    if ( day == 1 && evening ) {
      mnf = true;
      if ( mnfStartCol == undefined ) {
        mnfStartCol = headers.length + 1;
      }
      mnfEndCol = headers.length + 1;
    }
    let dayIndex = day + 3; // Day coloration function
    let writeCell = sheet.getRange(dayRow,firstMatchCol+(matches-1));
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false)))')
      .setBackground(dayColorsFilledObj[day])
      .setBold(true)
      .setRanges([writeCell]);
    rule.build();
    formatRules.push(rule);
    dayRowColors.push(dayColorsObj[day]);
    days.push(contests[a].dayName);
    const spread = contests[a].spread || '';
    spreads.push(spread);
    bonuses.push(contests[a].bonus);
    headers.push(away + '\n@' + home);
    widths.push(50);
    rule = SpreadsheetApp.newDataValidation().requireValueInList([away,home,"TIE"], true).build();
    sheet.getRange(outcomeRow,headers.length).setDataValidation(rule);
  }

  const finalMatchCol = headers.length;

  if (config.tiebreakerInclude) {
    headers.push('TIE\nBREAKER'); // Omitted if tiebreakers are removed
    widths.push(75);
    tiebreakerCol = headers.length;
    headers.push('DIFF');
    widths.push(50);
  }

  headers.push('WIN');
  widths.push(50);
  winCol = headers.indexOf('WIN')+1;

  if (!config.mnfExclude && mnf) {
    headers.push('MNF'); // Added if user wants a MNF competition included
    widths.push(50);
    mnfCol = headers.indexOf('MNF')+1;
  }

  if (!config.commentsExclude) {
    headers.push('COMMENT'); // Added to allow submissions to have amusing comments, if desired
    widths.push(150);
    commentCol = headers.indexOf('COMMENT')+1;
  }

  let diffCol = headers.length+1;
  let finalCol = diffCol + (diffCount-1);

  // Headers completed, now adjusting number of columns once headers are populated
  adjustColumns(sheet,finalCol);
  maxCols = sheet.getMaxColumns();

  sheet.getRange(matchRow,1,1,headers.length).setValues([headers]);
  sheet.getRange(dayRow,firstMatchCol,1,matches).setValues([days]);
 
  // spreadRow
  // outcomeRow
  // outcomeMarginRow
  // spreadOutcomeRow
  // bonusRow
  
  // Place spread values
  sheet.getRange(spreadRow,firstMatchCol,1,spreads.length).setValues([spreads])
  
  // Set Data validation for margin
  sheet.getRange(outcomeMarginRow,firstMatchCol,1,spreads.length).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Array.from({ length: 46 }, (_, index) => index), true).build());
  
  // Set Bonus values and validation
  let bonusRange = sheet.getRange(bonusRow,firstMatchCol,1,bonuses.length);
  bonusRange.setValues([bonuses]);
  rule = SpreadsheetApp.newDataValidation().requireValueInList(['1','2','3'],true).build();
  bonusRange.setDataValidation(rule);
  
  // Set all column widths
  for (let a = 0; a < widths.length; a++) {
    sheet.setColumnWidth(a+1,widths[a]);
  }
  

  // Begin building functions

  sheet.getRange(matchRow,diffCol).setValue('SIMILAR SELECTIONS'); // Added to allow submissions to have amusing comments, if desired
  sheet.getRange(dayRow,diffCol).setValue('Displayed as the number of picks different and the name of the member')
    .setFontSize(8);

  // Create named ranges
  ss.setNamedRange(`${LEAGUE}_${week}`,sheet.getRange(matchRow,firstMatchCol,1,matches)); // Then shortname versions of the matchups ( do have \n within )
  ss.setNamedRange(`${LEAGUE}_SPREADS_${week}`,sheet.getRange(spreadRow,firstMatchCol,1,matches)); // Spread values along bottom
  ss.setNamedRange(`${LEAGUE}_PICKEM_OUTCOMES_${week}`,sheet.getRange(outcomeRow,firstMatchCol,1,matches)); // Outcomes of game (straight up)
  ss.setNamedRange(`${LEAGUE}_PICKEM_OUTCOMES_${week}_MARGIN`,sheet.getRange(outcomeMarginRow,firstMatchCol,1,matches)); // Outcomes of game (straight up)
  ss.setNamedRange(`${LEAGUE}_ATS_OUTCOMES_${week}`,sheet.getRange(spreadOutcomeRow,firstMatchCol,1,matches)); // Outcomes of game (straight up)
  ss.setNamedRange(`${LEAGUE}_BONUS_${week}`,sheet.getRange(bonusRow,firstMatchCol,1,matches)); // Bonus multiplier for matchups
  ss.setNamedRange(`${LEAGUE}_PICKS_${week}`,sheet.getRange(entryRowStart,firstMatchCol,totalMembers,matches)); // All center data area (imported)

  if (!config.mnfExclude && mnf) {
    ss.setNamedRange(`${LEAGUE}_MNF_${week}`,sheet.getRange(entryRowStart,mnfStartCol,totalMembers,mnfEndCol-(mnfStartCol-1)));
  }
  if (config.tiebreakerInclude) {
    ss.setNamedRange(`${LEAGUE}_TIEBREAKER_${week}`,sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1));
    let validRule = SpreadsheetApp.newDataValidation().requireNumberBetween(0,120)
      .setHelpText('Must be a number')
      .build();
    sheet.getRange(outcomeRow,tiebreakerCol).setDataValidation(validRule);
  }
  if (!config.commentsExclude) {
    ss.setNamedRange(`COMMENTS_${week}`,sheet.getRange(entryRowStart,commentCol,totalMembers,1));
  }

  for (let row = entryRowStart; row <= entryRowEnd; row++ ) {
    // Formula to determine how many correct on the week
    sheet.getRange(row,1,1,maxCols).setBorder(null,null,true,null,false,false,'#AAAAAA',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    sheet.getRange(row,pointsCol).setFormulaR1C1('=iferror(if(and(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+')>0,counta(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')>0),sum(arrayformula(if(not(isblank(R'+row+'C'+firstMatchCol+':R'+row+'C'+finalMatchCol+')),if(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+'=R'+row+'C'+firstMatchCol+':R'+row+'C'+finalMatchCol+',1,0),0)*R'+bonusRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol+')),))');

    // sheet.getRange(row,2).setFormulaR1C1('=iferror(if(and(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C['+finalMatchCol+'])>0,counta(R[0]C[3]:R[0]C['+finalMatchCol+'])>0),mmult(arrayformula(if(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+'=R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')),1,0)))),))');
    
    // Formula to determine weekly rank
    sheet.getRange(row,pointsCol+1).setFormulaR1C1('=iferror(if(and(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+')>0,not(isblank(R[0]C'+pointsCol+'))),rank(R[0]C'+pointsCol+',R'+entryRowStart+'C2:R'+entryRowEnd+'C2,false),))');

    // Formula to determine weekly correct percent
    sheet.getRange(row,pointsCol+2).setFormulaR1C1('=iferror(if(and(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+')>0,not(isblank(R[0]C'+pointsCol+'))),sum(filter(arrayformula(if(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+'=R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+',1,0)),not(isblank(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+'))))/counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+'),),)');
    
    // Formula to determine difference of tiebreaker from final MNF score
    if (config.tiebreakerInclude) {
      sheet.getRange(row,tiebreakerCol+1).setFormulaR1C1('=iferror(if(or(isblank(R[0]C[-1]),isblank(R'+outcomeRow+'C'+tiebreakerCol+')),,abs(R[0]C[-1]-R'+outcomeRow+'C'+tiebreakerCol+')))');
      // Formula to denote winner with a '1' if a clear winner exists
      sheet.getRange(row,winCol).setFormulaR1C1('=iferror(if(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')=value(regexextract(R'+dayRow+'C1,\"[0-9]+\")),arrayformula(if(countif(array_constrain({R[0]C'+pointsCol+',R[0]C'+(tiebreakerCol+1)+'}=filter(filter({R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+',R'+entryRowStart+'C'+(tiebreakerCol+1)+':R'+entryRowEnd+'C'+(tiebreakerCol+1)+'},R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+'=max(R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+')),filter(R'+entryRowStart+'C'+(tiebreakerCol+1)+':R'+entryRowEnd+'C'+(tiebreakerCol+1)+',R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+'=max(R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+'))=min(filter(R'+entryRowStart+'C'+(tiebreakerCol+1)+':R'+entryRowEnd+'C'+(tiebreakerCol+1)+',R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+'=max(R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+')))),1,2),true)=2,1,0))),)');
    } else {
      // Formula to denote winner with a '1', with a tiebreaker allowed
      sheet.getRange(row,winCol).setFormulaR1C1('=iferror(if(counta(R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+')=value(regexextract(R'+dayRow+'C1,\"[0-9]+\")),if(rank(R'+row+'C'+pointsCol+',R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol+',false)=1,1,0)),)');
    }

    // Formula to determine MNF win status sum (can be more than 1 for rare weeks)
    if (config.mnfInclude && mnf) {
      sheet.getRange(row,mnfCol).setFormulaR1C1('=iferror(if(and(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+')>0,counta(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')>0),if(mmult(arrayformula(if(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfStartCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfEndCol+'=R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+')),1,0))))=0,0,mmult(arrayformula(if(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfStartCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfEndCol+'=R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+',1,0)),transpose(arrayformula(if(not(isblank(R[0]C'+mnfStartCol+':R[0]C'+mnfEndCol+')),1,0))))),),)');
    }

    // Formula to generate array of similar pickers on the week
    sheet.getRange(row,diffCol).setFormulaR1C1('=iferror(if(isblank(R[0]C'+(firstMatchCol+2)+'),,transpose(arrayformula({arrayformula('+matches+'-query({R'+entryRowStart+'C1:R'+entryRowEnd+'C1,arrayformula(mmult(if(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+(finalMatchCol)+'=R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+',1,0),transpose(arrayformula(column(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')\^0))))},\"select Col2 where Col1 <> \'\"\&R[0]C1\&\"\' order by Col2 desc, Col1 asc limit '+diffCount+'\"))&\": \"&query({R'+entryRowStart+'C1:R'+entryRowEnd+'C1,arrayformula(mmult(if(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+'=R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+',1,0),transpose(arrayformula(column(R[0]C'+firstMatchCol+':R[0]C'+finalMatchCol+')\^0))))},\"select Col1 where Col1 \<\> \'\"\&R[0]C1\&\"\' order by Col2 desc, Col1 asc limit '+diffCount+
      '\")}))))');
  }

  // Sets the formula for home / away split for each matchup column
  for (let col = firstMatchCol; col <= finalMatchCol; col++ ) {
    sheet.getRange(summaryRow,col).setFormulaR1C1('=iferror(if(counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])>0,if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],regexextract(R'+matchRow+'C[0],"[A-Z]{2,3}"))=counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])/2,\"SPLIT\"&char(10)&\"50%\",if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],regexextract(R'+matchRow+'C[0],\"[A-Z]{2,3}\"))<counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])/2,regexextract(right(R'+matchRow+'C[0],3),\"[A-Z]{2,3}\")&char(10)&round(100*countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],regexextract(right(R'+matchRow+'C[0],3),\"[A-Z]{2,3}\"))/counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1)&\"%\",regexextract(R'+matchRow+'C[0],\"[A-Z]{2,3}\")&char(10)&round(100*countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],regexextract(R'+matchRow+'C[0],\"[A-Z]{2,3}\"))/counta(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1)&\"%\")),))');

    sheet.getRange(spreadOutcomeRow,col).setFormulaR1C1(`=iferror(if(and(not(isblank(R[-2]C[0])),not(isblank(R[-1]C[0]))),if(or(R[-2]C[0]="TIE",not(regexextract(R[-3]C[0],"[A-Z]{2,3}")=R[-2]C[0])),trim(regexreplace(regexreplace(R1C[0],regexextract(R[-3]C[0],"[A-Z]{2,3}"),""),"@","")),if(and(regexextract(R[-3]C[0],"[A-Z]{2,3}")=R[-2]C[0],R[-1]C[0]>-value(regexextract(R[-3]C[0],"[0-9\-]+"))),R[-2]C[0],trim(regexreplace(regexreplace(R1C[0],R[-2]C[0],""),"@","")))),))`);
  }
  
  if (config.tiebreakerInclude) {
    sheet.getRange(matchRow,winCol).setFormulaR1C1('=if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)>1,\"TIE\",\"WIN\")');
    sheet.getRange(summaryRow,winCol).setFormulaR1C1('=iferror(if(not(isblank(R'+summaryRow+'C'+tiebreakerCol+')),if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)>1,countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)&\"\-WAY\"&char(10)&\"TIE\",),),)');
    sheet.getRange(summaryRow,tiebreakerCol).setFormulaR1C1('=iferror(if(sum(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])>0,\"AVG\"&char(10)&round(average(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1),),)');
    sheet.getRange(summaryRow,tiebreakerCol+1).setFormulaR1C1('=iferror(if(sum(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])>0,\"AVG\"&char(10)&round(average(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1),),)');
  } else {
    sheet.getRange(summaryRow,winCol).setFormulaR1C1('=iferror(if(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+')=value(regexextract(R'+dayRow+'C1,\"[0-9]+\")),if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)>1,countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)&\"\-WAY\"&char(10)&\"TIE\",\"DONE\"),),)');
    sheet.getRange(matchRow,winCol).setFormulaR1C1('=iferror(if(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+')=value(regexextract(R'+dayRow+'C1,\"[0-9]+\")),if(countif(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0],1)=0,\"TIE\",\"WIN\"),\"WIN\"),)');
  }

  if (config.mnfInclude && mnf) {
    sheet.getRange(summaryRow,mnfCol).setFormulaR1C1('=iferror(if(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfStartCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfEndCol+')=columns(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfStartCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfEndCol+'),\"MNF\"\&char(10)&(round(sum(mmult(arrayformula(if(R'+entryRowStart+'C'+mnfStartCol+':R'+entryRowEnd+'C'+mnfEndCol+'=R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfStartCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfEndCol+',1,0)),transpose(arrayformula(if(not(isblank(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfStartCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+mnfEndCol+')),1,0)))))/counta(R'+entryRowStart+'C'+mnfStartCol+':R'+entryRowEnd+'C'+mnfEndCol+'),3)*100)\&\"\%\",),)');
  }

  sheet.getRange(matchRow,pointsCol).setFormulaR1C1('=iferror(if(countif(R'+bonusRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol+',\">1\")>0,\"TOTAL\"&char(10)&\"POINTS\",\"TOTAL\"&char(10)&\"CORRECT\"),)');

  sheet.getRange(summaryRow,pointsCol).setFormulaR1C1('=iferror(if(sum(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0])>0,\"AVG\"\&char(10)&(round(average(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),1)),),)');

  sheet.getRange(summaryRow,diffCol).setFormulaR1C1('=iferror(if(isblank(R[0]C'+firstMatchCol+'),,transpose(query({arrayformula((counta(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+')-mmult(arrayformula(if(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+'=arrayformula(regexextract(R'+(totalMembers+3)+'C'+firstMatchCol+':R'+(totalMembers+3)+'C'+finalMatchCol+',\"[A-Z]+\")),1,0)),transpose(arrayformula(if(arrayformula(len(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+'))>1,1,1)))))&\": \"\&'+'R'+entryRowStart+'C1:R'+entryRowEnd+'C1),mmult(arrayformula(if(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+'=arrayformula(regexextract(R'+(totalMembers+3)+'C'+firstMatchCol+':R'+(totalMembers+3)+'C'+finalMatchCol+',\"[A-Z]+\")),1,0)),transpose(arrayformula(if(arrayformula(len(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+'))>1,1,1))))},\"select Col1 order by Col2 desc, Col1 desc limit '+diffCount+'\"))))');

  // AWAY TEAM BIAS FORMULA 
  sheet.getRange(summaryRow,2,1,1).setFormulaR1C1('=iferror(if(counta(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+')>10,"AWAY"&char(10)&round(100*(sum(arrayformula(if(regexextract(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+',"^[A-Z]{2,3}")=R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+',1,0)))/counta(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+')),1)&"%","AWAY"),"AWAY")');
  // HOME TEAM BIAS FORMULA
  sheet.getRange(summaryRow,3,1,1).setFormulaR1C1('=iferror(if(counta(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+')>10,"HOME"&char(10)&round(100*(sum(arrayformula(if(regexextract(R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol+',"[A-Z]{2,3}$")=R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+',1,0)))/counta(R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol+')),1)&"%","HOME"),"HOME")');
  sheet.getRange(summaryRow,4,1,1).setFormulaR1C1('=iferror(if(counta(R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+firstMatchCol+':R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C'+finalMatchCol+')>2,average(R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]),))');

  // Setting conditional formatting rules
  let bonusCount = 3;
  let parity = ['iseven','isodd'];
  let formatObj = [{'name':'correct_pick_even','color_start':'#c9ffdf','color_end':'#69ffa6','formula':'=and(indirect(\"R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C[0]\",false))),'+parity[0]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'correct_pick_odd','color_start':'#a0fdba','color_end':'#73ff9b','formula':'=and(indirect(\"R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C[0]\",false)=indirect(\"R[0]C[0]\",false),not(isblank(indirect(\"R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C[0]\",false))),'+parity[1]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'incorrect_pick_even','color_start':'#FFF7F9','color_end':'#FCD4DC','formula':'=and(not(isblank(indirect(\"R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C[0]\",false))),'+parity[0]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'incorrect_pick_odd','color_start':'#FFF2F4','color_end':'#FFC3CC','formula':'=and(not(isblank(indirect(\"R'+(isAts ? spreadOutcomeRow : outcomeRow)+'C[0]\",false))),'+parity[1]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'home_pick_even','color_start':'#e3fffe','color_end':'#9ef2ee','formula':'=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),arrayformula(trim(split(indirect(\"R'+matchRow+'C[0]\",false),\"\@\"))),0)=2,'+parity[0]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'home_pick_odd','color_start':'#d0f5f3','color_end':'#80f1ea', 'formula':'=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),arrayformula(trim(split(indirect(\"R'+matchRow+'C[0]\",false),\"\@\"))),0)=2,'+parity[1]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'away_pick_even','color_start':'#fffee3','color_end':'#fdf9a2','formula':'=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),arrayformula(trim(split(indirect(\"R'+matchRow+'C[0]\",false),\"\@\"))),0)=1,'+parity[0]+'(row(indirect(\"R[0]C1\",false))))'},
                {'name':'away_pick_odd','color_start':'#faf9e1','color_end':'#fbf77f','formula':'=and(not(isblank(indirect(\"R[0]C[0]\",false))),match(indirect(\"R[0]C[0]\",false),arrayformula(trim(split(indirect(\"R'+matchRow+'C[0]\",false),\"\@\"))),0)=1,'+parity[1]+'(row(indirect(\"R[0]C1\",false))))'}];

  sheet.clearConditionalFormatRules();    
  let range = sheet.getRange('R'+entryRowStart+'C'+firstMatchCol+':R'+entryRowEnd+'C'+finalMatchCol);
  Object.keys(formatObj).forEach(a => {
    let gradient = hexGradient(formatObj[a].color_start,formatObj[a].color_end,bonusCount);
    for (let b = gradient.length-1; b >= 0; b--) {
      let formula = formatObj[a].formula;
      if (b > 0) {
        // Appends the number bonus amount to the conditional formatting to pair with the gradient value assigned
        formula = formula.substring(0,formula.length-1).concat(',indirect(\"R'+bonusRow+'C[0]\",false)='+(b+1)+')');
      }
      let rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(gradient[b])
        .setRanges([range]);
      if (formatObj[a].name.includes('incorrect')) {
        rule.setFontColor('#999999'); // Dark gray text for the incorrect picks
      }
      rule.build();
      
      formatRules.push(rule);
    }
  });

  // NAMES COLUMN NAMED RANGE
  range = sheet.getRange('R'+entryRowStart+'C1:R'+entryRowEnd+'C1');
  ss.setNamedRange('NAMES_'+week,range);

  // TOTALS GRADIENT RULE
  range = sheet.getRange('R'+entryRowStart+'C2:R'+entryRowEnd+'C2');
  ss.setNamedRange('TOT_'+week,range);
  let formatRuleTotals = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpoint("#75F0A1")
    .setGradientMinpoint("#FFFFFF")
    //.setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, (finalMatchCol-2) - 3) // Max value of all correct picks (adjusted by 3 to tighten color range)
    //.setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, (finalMatchCol-2) / 2)  // Generates Median Value
    //.setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, 0 + 3) // Min value of all correct picks (adjusted by 3 to tighten color range)
    .setRanges([range])
    .build();
  formatRules.push(formatRuleTotals);
  // RANKS GRADIENT RULE
  range = sheet.getRange('R'+entryRowStart+'C3:R'+entryRowEnd+'C3');
  ss.setNamedRange('RNK_'+week,range);
  let formatRuleRanks = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, members.length)
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, members.length/2)
    .setGradientMinpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([range])
    .build();
  formatRules.push(formatRuleRanks);
  // PERCENT GRADIENT RULE
  range = sheet.getRange('R'+entryRowStart+'C4:R'+(rows)+'C4');
  range.setNumberFormat('##0.0%');
  let formatRulePercent = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#75F0A1", SpreadsheetApp.InterpolationType.NUMBER, ".70")
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, ".60")
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, ".50")
    .setRanges([range])
    .build();
  formatRules.push(formatRulePercent);
  ss.setNamedRange('PCT_'+week,sheet.getRange('R'+entryRowStart+'C4:R'+entryRowEnd+'C4'));    
  // POINTS GRADIENT RULE
  range = sheet.getRange('R'+entryRowStart+'C'+pointsCol+':R'+entryRowEnd+'C'+pointsCol);
  let formatRulePoints = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue("#5EDCFF", SpreadsheetApp.InterpolationType.NUMBER, '=max(indirect(\"R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]\",false))')
    .setGradientMidpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, '=average(indirect(\"R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]\",false))')
    .setGradientMinpointWithValue("#FF9B69", SpreadsheetApp.InterpolationType.NUMBER, '=min(indirect(\"R'+entryRowStart+'C[0]:R'+entryRowEnd+'C[0]\",false))')
    .setRanges([range])
    .build();
  formatRules.push(formatRulePoints);


  // WINNER COLUMN RULE
  range = sheet.getRange('R'+entryRowStart+'C'+winCol+':R'+entryRowEnd+'C'+winCol);
  ss.setNamedRange('WIN_'+week,range);
  let formatRuleNotWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberNotEqualTo(1)
    .setBackground('#FFFFFF')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build();     
  formatRules.push(formatRuleNotWinner);
  let formatRuleWinner = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(1)
    .setBackground('#75F0A1')
    .setFontColor('#75F0A1')
    .setRanges([range])
    .build();
  formatRules.push(formatRuleWinner);  
  // WINNER NAME RULE
  range = sheet.getRange('R'+entryRowStart+'C1:R'+entryRowEnd+'C1');
  let formatRuleWinnerName = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=indirect(\"R[0]C'+winCol+'\",false)=1')
    .setBackground('#75F0A1')
    .setRanges([range])
    .build();
  formatRules.push(formatRuleWinnerName);

  // MNF GRADIENT RULE
  let formatRuleMNFEmpty, formatRuleMNF;
  if (config.mnfInclude && mnf) {
    range = sheet.getRange('R'+entryRowStart+'C'+mnfCol+':R'+entryRowEnd+'C'+mnfCol);
    ss.setNamedRange('MNF_'+week,range);
    formatRuleMNFEmpty = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=or(isblank(indirect("R[0]C[0]",false)),indirect("R[0]C[0]",false)=0)')
      .setFontColor('#FFFFFF')
      .setBackground('#FFFFFF')
      .setRanges([range])
      .build();
    formatRules.push(formatRuleMNFEmpty);      
    if (mnfStartCol != mnfEndCol) { // Rules for when there are multiple MNF games
      formatRuleMNF = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpoint("#FFF624") // Max value of all correct picks, min 1
        .setGradientMinpoint("#FFFFFF") // Min value of all correct picks  
        .setRanges([range])
        .build();
    } else { // Rules for single MNF game 
      formatRuleMNF = SpreadsheetApp.newConditionalFormatRule()
        .setBackground("#FFF624")
        .setFontColor("#FFF624")
        .whenNumberEqualTo(1)
        .setRanges([range])
        .build();
    }
    formatRules.push(formatRuleMNF);
  }
 
  // DIFFERENCE TIEBREAKER COLUMN FORMATTING
  if (config.tiebreakerInclude) {
    let offsets = [1,3,5,10,15,20,20];
    let offsetColors = hexGradient('#33FF7A','#FFFFFF',offsets.length);
    for (let a = 0; a < offsets.length; a++) {
      let rule;
      if (a < (offsets.length - 1)) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),abs(indirect(\"R[0]C[0]\",false)-indirect(\"R'+outcomeRow+'C[0]:R'+outcomeRow+'C[0]\",false))<='+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1)])
          .build();
      } else {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),abs(indirect(\"R[0]C[0]\",false)-indirect(\"R'+outcomeRow+'C[0]:R'+outcomeRow+'C[0]\",false))>'+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1)])
          .build();        
      }
      formatRules.push(rule);
      rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[0]\",false))),abs(value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))-indirect(\"R'+outcomeRow+'C[0]:R'+outcomeRow+'C[0]\",false))<='+offsets[a]+',)')
        .setBackground(offsetColors[a])
        .setRanges([sheet.getRange(summaryRow,tiebreakerCol)])
        .build();
      formatRules.push(rule);
    }
    offsetColors = hexGradient('#FFFFFF','#666666',offsets.length);
    for (let a = 0; a < offsets.length; a++) {
      let rule;
      let ruleOffsets;
      if (a < (offsets.length - 1)) {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[-1]\",false))),indirect(\"R[0]C[0]\",false)<='+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(entryRowStart,tiebreakerCol+1,totalMembers,1)])
          .build();
        ruleOffsets = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[-1]\",false))),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))<='+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(summaryRow,tiebreakerCol+1)])
          .build();
      } else {
        rule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[-1]\",false))),indirect(\"R[0]C[0]\",false)>'+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(entryRowStart,tiebreakerCol+1,totalMembers,1)])
          .build();
        ruleOffsets = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied('=if(not(isblank(indirect(\"R'+outcomeRow+'C[-1]\",false))),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))>'+offsets[a]+',)')
          .setBackground(offsetColors[a])
          .setRanges([sheet.getRange(summaryRow,tiebreakerCol+1)])
          .build();              
      }
      formatRules.push(rule);
      formatRules.push(ruleOffsets);
    }
    // ADD ADDITIONAL COLOR VARIATION BASED ON TIEBREAKER VALUE PRESENT HERE
    let formatRuleTiebreakerEmptyAndDone = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=and(isblank(indirect(\"R[0]C[0]\",false)),counta(indirect(\"R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+'\",false))>=columns(indirect(\"R'+outcomeRow+'C'+firstMatchCol+':R'+outcomeRow+'C'+finalMatchCol+'\",false)))')
      .setBackground("#FF3FC7")
      .setRanges([sheet.getRange(outcomeRow,tiebreakerCol)])
      .build();
    formatRules.push(formatRuleTiebreakerEmptyAndDone);
    let formatRuleTiebreakerEmpty = SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground("#CCCCCC")
      .setRanges([sheet.getRange(outcomeRow,tiebreakerCol)])
      .build();
    formatRules.push(formatRuleTiebreakerEmpty);
    range = sheet.getRange(entryRowStart,tiebreakerCol,totalMembers,1);
    let formatRuleDiff = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpoint("#B7B7B7")
      .setGradientMinpoint("#FFFFFF")
      .setRanges([range])
      .build();
    formatRules.push(formatRuleDiff);
  }

  // PREFERENCE COLOR SCHEMES
  let homeAwayPercents = [90,80,70,60,50];
  let awayColors = ['#FFFB7D','#FFFC96','#FFFCB0','#FFFDC9','#FFFEE3'];
  let homeColors = ['#7DFFFB','#96FFFC','#B0FFFC','#C9FFFD','#E3FFFE'];
  let awayFormula = '=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(indirect(\"R'+matchRow+'C[0]\",false),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=%%)'; // Replaceable "%%" for inserting percent number
  let homeFormula = '=and(regexextract(indirect(\"R[0]C[0]\",false),\"[A-Z]{2,3}\")=regexextract(right(indirect(\"R'+matchRow+'C[0]\",false),3),\"[A-Z]{2,3}\"),value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9\\.]+\"))>=%%)'; // Replaceable "%%" for inserting percent number
  range = sheet.getRange(summaryRow,firstMatchCol,1,matches); // Summary row of matches
  for (let a = 0; a < homeAwayPercents.length; a++) {
    let formula = awayFormula.replace('%%',homeAwayPercents[a]);

    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(awayColors[a])
      .setRanges([range]);
    rule.build();
    formatRules.push(rule);

    formula = homeFormula.replace('%%',homeAwayPercents[a]);
    rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(homeColors[a])
      .setRanges([range]);
    rule.build();
    formatRules.push(rule);    
  }

  // MATCHUP WEIGHTING RULE
  let formatRuleWeightedThree, formatRuleWeightedTwo;
  formatRuleWeightedThree = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),or(and(indirect(\"R'+bonusRow+'C[0]\",false)=2,countif(indirect(\"R'+bonusRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol+'\",false),3)=0),indirect(\"R'+bonusRow+'C[0]\",false)=3))')
    .setBackground('#9C9C97')
    .setRanges([sheet.getRange('R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol),sheet.getRange('R'+spreadRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol)])
    .build();
  formatRules.push(formatRuleWeightedThree);
  formatRuleWeightedTwo = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=and(not(isblank(indirect(\"R[0]C[0]\",false))),indirect(\"R'+bonusRow+'C[0]\",false)=2)')
    .setBackground('#949376')
    .setRanges([sheet.getRange('R'+matchRow+'C'+firstMatchCol+':R'+matchRow+'C'+finalMatchCol),sheet.getRange('R'+spreadRow+'C'+firstMatchCol+':R'+bonusRow+'C'+finalMatchCol)])
    .build();
  formatRules.push(formatRuleWeightedTwo);
  
  // Format rules for difference columns to emphasize the most common picker
  let commonPickersGradient = hexGradient('#46f081','#e4f0e8',8);
  let commonPickersFormula = '=value(regexextract(indirect(\"R[0]C[0]\",false),\"[0-9]+\"))=%'; // Replaceable "%" for common picker number
  range = sheet.getRange(entryRowStart,diffCol,totalMembers+1,diffCount);
  for (let a = 0; a < commonPickersGradient.length; a++) {
    let formula = commonPickersFormula.replace('%',a); // Replaces "%" with index of commonPickersGradient (0-X)
    let rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(commonPickersGradient[a])
      .setRanges([range])
      .build();
    formatRules.push(rule);
  }

  // Sets all formerly pushed rules to the sheet
  sheet.setConditionalFormatRules(formatRules);

  // Setting size, alignment, frozen columns
  columns = sheet.getMaxColumns();
  sheet.getRange(1,1,rows,columns)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setFontSize(10)
    .setFontFamily("Montserrat");

  sheet.getRange(entryRowStart,diffCol,totalMembers+1,diffCount).setHorizontalAlignment('left');
  if (!config.commentsExclude) {
    sheet.getRange(2,commentCol,totalMembers+1,1).setHorizontalAlignment('left');
  }

  sheet.getRange(1,1,summaryRow,1)
    .setHorizontalAlignment('left');
 
  sheet.setFrozenColumns(firstMatchCol-1);
  sheet.setFrozenRows(dayRow);
  sheet.getRange(1,1,1,columns)
    .setBackground('black')
    .setFontColor('white')
    .setFontWeight('bold');
  sheet.setRowHeights(1,rows,21);

  sheet.getRange(matchRow,1,1,sheet.getMaxColumns()).setVerticalAlignment('middle');
  sheet.setRowHeight(matchRow,50);
  sheet.getRange(matchRow,1).setFontSize(18)
    .setHorizontalAlignment('center');
  
  sheet.getRange(dayRow,firstMatchCol,1,matches).setFontSize(7);
  sheet.getRange(dayRow,1,1,maxCols).setBackground('#CCCCCC');
  sheet.getRange(dayRow,firstMatchCol,1,dayRowColors.length).setBackgrounds([dayRowColors]);
  sheet.getRange(dayRow,1,1,firstMatchCol-1).mergeAcross();
  sheet.getRange(dayRow,1).setValue(matches + ' ' + LEAGUE + ' MATCHES')
    .setHorizontalAlignment('left');
  
  // Spread row
  sheet.getRange(spreadRow,1,1,firstMatchCol-1).mergeAcross().setHorizontalAlignment('right');  
  sheet.getRange(spreadRow,1,1,maxCols).setBackground('black')
    .setFontColor('white')
    .setFontSize(8);
  // Outcome row
  sheet.getRange(outcomeRow,1,1,firstMatchCol-1).mergeAcross().setHorizontalAlignment('right');  
  sheet.getRange(outcomeRow,1,1,maxCols).setBackground('black')
    .setFontColor('white')
    .setFontWeight('bold');
  // Outcome Margin row
  sheet.getRange(outcomeMarginRow,1,1,firstMatchCol-1).mergeAcross().setHorizontalAlignment('right');  
  sheet.getRange(outcomeMarginRow,1,1,maxCols).setBackground('black')
    .setFontColor('white');
  // Spread Winner row
  sheet.getRange(spreadOutcomeRow,1,1,firstMatchCol-1).mergeAcross().setHorizontalAlignment('right');  
  sheet.getRange(spreadOutcomeRow,1,1,maxCols).setBackground('black')
    .setFontColor('white')
    .setFontWeight('bold');        
  // Bonus Row
  sheet.getRange(bonusRow,1,1,firstMatchCol-1).mergeAcross().setHorizontalAlignment('right');    
  sheet.getRange(bonusRow,1,1,maxCols).setBackground('black')
    .setFontColor('white');
  if (!config.bonusInclude) {
    sheet.hideRows(bonusRow);
  }
  if (!isAts) {
    sheet.hideRows(spreadRow);
    sheet.hideRows(outcomeMarginRow);
    sheet.hideRows(spreadOutcomeRow);
  }

  sheet.setRowHeight(summaryRow,40);
  sheet.getRange(summaryRow,1,1,sheet.getMaxColumns()).setVerticalAlignment('middle');
  sheet.getRange(summaryRow,1,1,maxCols-diffCount).setBackground('#CCCCCC');
  sheet.getRange(summaryRow,2).setBackground(awayColors[1]);
  sheet.getRange(summaryRow,3).setBackground(homeColors[1]);

  sheet.setColumnWidths(diffCol,diffCount,90);
  sheet.getRange(1,diffCol,2,diffCount)
    .setHorizontalAlignment('left')
    .mergeAcross();

  if (config.tiebreakerInclude) {
    sheet.getRange(outcomeRow,tiebreakerCol).setNote('Enter the summed score of the outcome of the final game of the week in this cell to complete the week and designate a winner');
  }
  sheet.getRange(dayRow,finalMatchCol+1,1,finalCol-finalMatchCol-diffCount).mergeAcross();

  const text = `✅ Completed creation of pick 'ems week ${week} sheet.`;
  Logger.log(text)
  ss.toast(text,'SUCCESS');
  return sheet;
}

// WEEKLY SHEET COLORATION - Adds a color to the weekly tabs that exist and uses the "dayColorsFilled" array [global variable]
function weeklySheetTabColors(ss,sheet) {
  let week;
  ss = fetchSpreadsheet(ss);
  if (sheet == undefined) {
    week = ss.getRangeByName('WEEK').getValue();
    sheet = ss.getSheetByName(weeklySheetPrefix + week);
  }
  try {
    if (sheet == undefined) {
      throw new Error();
    }
    let colors = [...dayColorsFilled];
    colors.push(winnersTabColor); // Adds a bright yellow to the end of the array for the active week tab
    let week = parseInt(sheet.getName().replace(weeklySheetPrefix,''));
    sheet.setTabColor(colors[colors.length-1]);
    colors.pop();
    for (let a = (week - 1); a > 0; a--) {
      let sheet = ss.getSheetByName(weeklySheetPrefix + a);
      if (sheet != null) {
        sheet.setTabColor(colors[colors.length-1]);
      }
      if (colors.length > 1) {
        colors.pop();
      }
    }
    Logger.log('changed all colors of tabs to reflect week shift');
  }
  catch (err) {
    Logger.log('Error assigning colors to weekly sheet tabs: ' + err.stack);
  }
}


// ============================================================================================================================================
// BONUS FEATURES
// ============================================================================================================================================

// GAME OF THE WEEK SHEET FUNCTION - selects one random game for 2x multiplier to be applied
function bonusRandomGameSet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let tnf = true, bonusRange, mnfDouble = false, text;
  const week = maxWeek();
  let sheet, sheetExisted = true;
  try { 
    ss.getRangeByName('BONUS_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('No \'BONUS_PRESENT\' named range');
    ui.alert('BONUS PRESENT NOT SET\r\n\r\nNo bonus present range established for inclusion/exclusion of bonus game weighting, please run the enable/disable bonus function and try this function again after that has been set', ui.ButtonSet.OK);
    throw new Error('Canceled due to no bonus feature');
  }
  try { 
    tnf = ss.getRangeByName('TNF_PRESENT').getValue();
  }
  catch (err) {
    Logger.log('No \'TNF_PRESENT\' named range, assuming true');
    ui.alert('THURSDAY NIGHT FOOTBALL EXCLUSION NOT SET\r\n\r\nNo Thursday present range established for inclusion/exclusion of Thursday\'s games', ui.ButtonSet.OK);
  }
  try { 
    mnfDouble = ss.getRangeByName('MNF_DOUBLE').getValue();
  }
  catch (err) {
    Logger.log('No \'MNF_DOUBLE\' named range, assuming \'false\' for double MNF and proceeding.');
  }
  try {
    sheet = ss.getSheetByName(weeklySheetPrefix + week);
  } catch (err) {
    Logger.log('No sheet for week ' + week);
    let prompt = ui.alert('NO SHEET\r\n\r\nThe week ' + week + ' sheet does not exist. Create a sheet now for week ' + week + '?\r\n\r\n(Selecting \'Cancel\' will exit and no game will be selected)', ui.ButtonSet.OK_CANCEL);
    if (prompt == ui.Button.OK) {
      sheet = weeklySheet(ss,week,config,members,false);
      sheetExisted = false;
    } else {
      throw new Error('Exited when new sheet creation was declined');
    }
  }
  try {
    bonusRange = ss.getRangeByName(LEAGUE + '_BONUS_' + week);
  }
  catch (err) {
    Logger.log('No \'BONUS\' named range for week ' + week);
    ui.alert('NO BONUS\r\n\r\nThe week ' + week + ' sheet lacks the bonus game feature. Would you like to recreate the week ' + week + ' sheet now?\r\n\r\n(Selecting \'Cancel\' will exit and no game will be selected)', ui.ButtonSet.OK_CANCEL);
    if (prompt == ui.Button.OK) {
      sheet = weeklySheet(ss,week,config,members,false);
      sheetExisted = false;
    } else {
      throw new Error('Exited when new sheet creation was declined');
    }
  }
  bonusRange = ss.getRangeByName(LEAGUE + '_BONUS_' + week);

  let mnf = false, mnfRange, bonusValues = bonusRange.getValues().flat();

  if (mnfDouble) {
    try {
      mnfRange = ss.getRangeByName(LEAGUE + '_MNF_' + week);
      bonusValues.splice(bonusValues.length-mnfRange.getNumColumns(),mnfRange.getNumColumns());
      bonusRange = sheet.getRange(bonusRange.getRow(),bonusRange.getColumn(),1,bonusRange.getNumColumns()-mnfRange.getNumColumns());
      if (mnfRange.getValues().length > 0) {
        mnf = true;
      }
    }
    catch (err) {
      Logger.log('No MNF range for week ' + week + '. Including all games in randomization.');
    }
  }  

  if (sheetExisted) {
    for (let a = 0; a < bonusValues.length; a++) {
      if (bonusValues[a] > 1) {
        text = 'BONUS GAME ALREADY MARKED\r\n\r\nYou already have one or more games marked for 2x or greater weighting.\r\n\r\nMark all ';
        if (mnfDouble && mnf) {
          text = text.concat('non-MNF games\' weighting to 1 and try again');
        } else {
          text = text.concat('games\' weighting to 1 and try again');
        }
        ui.alert(text,ui.ButtonSet.OK);
        throw new Error('Other games marked as bonus prior to running random Game of the Week function');
      }
    }
  }

  text = 'GAME OF THE WEEK\r\n\r\nWould you like to randomly select one game this week to count as double?';
  if (mnfDouble) {
    text = text.concat('\r\n\r\nAny MNF games will be excluded since you have the MNF Double feature enabled');
  }
  let gameOfTheWeek;
  let randomPrompt = ui.alert(text,ui.ButtonSet.YES_NO);
  if (randomPrompt == ui.Button.YES) {
    gameOfTheWeek = bonusRandomGame(week,tnf,mnfDouble);
    let matchupNames = ss.getRangeByName(LEAGUE + '_' + week).getValues().flat();
    let regex = new RegExp(/[A-Z]{2,3}/,'g');
    let matchupRegex = [];
    matchupNames.forEach(a => matchupRegex.push(a.match(regex)[0]+ '@' + a.match(regex)[1]));
    bonusValues[matchupRegex.indexOf(gameOfTheWeek)] = 2;
    bonusRange.setValues([bonusValues]);
  }

  let formId = ss.getRangeByName('FORM_WEEK_' + week).getValue();
  try {
    let form = FormApp.openById(formId);
    let prompt = ui.alert('FORM EXISTS\r\n\r\nYou\'ve already created a form for week ' + week + ', would you like to designate the Game of the Week on the Form?',ui.ButtonSet.YES_NO);
    if (prompt == ui.Button.YES) {
      let form = FormApp.openById(formId);
      let questions = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
      for (let a = 0; a < questions.length; a++) {
        try{
          let choices = questions[a].asMultipleChoiceItem().getChoices();
          let matchup = choices[0].getValue() + '@' + choices[1].getValue();
          if (matchup == gameOfTheWeek) {
            questions[a].setTitle('GAME OF THE WEEK (Double Points)\n' + questions[a].getTitle());
            break;
          }
        }
        catch (err) {
          Logger.log('Issue with getting choices for question with title ' + questions[a].getTitle() + ' or setting the title.');
        }
      }
    }
  }
  catch (err) {
    Logger.log('No form exists for week ' + week + ' or there was an error getting the questions for the form.'); 
  }
}

// GAME OF THE WEEK SELECTION - selects one random game for 2x multiplier to be applied
function bonusRandomGame(week,tnf,mnfDouble) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (week == null) {
    week = maxWeek();
  }

  let games = fetchGames(week);
  
  let abbrevs = [];
  let teams = [];
  for (let a = games.length - 1; a >= 0; a--) {
    if ((games[a][1] == 1 && mnfDouble) || (games[a][1] == -3 && !tnf)) {
      games.splice(a,1);
    } else {
      abbrevs.push(games[a][5] + '@' + games[a][6]);
      teams.push(games[a][7] + ' ' + games[a][8] + ' at ' + games[a][9] + ' ' + games[a][10]);
    }
  }

  let gameOfTheWeekIndex = getRandomInt(0,abbrevs.length-1);

  text = 'For week ' + week + ', your Game of the Week has been randomly selected as:\r\n\r\n';
  try {
    let gameOfTheWeek = abbrevs[gameOfTheWeekIndex];
    text = text.concat(teams[gameOfTheWeekIndex] + '\r\n\r\nWould you like to mark it as such?');
    let verify = ui.alert(text,ui.ButtonSet.OK_CANCEL);
    if (verify == ui.Button.OK) {
      return gameOfTheWeek;
    } else {
      ss.toast('Canceled Game of the Week selection');
    }
  }
  catch (err) {
    ss.toast('Error fetching matches or selecting Game of the Week\r\n\r\nError:\r\n' + err.message);
    Logger.log('Error fetching matches or selecting Game of the Week\r\n\r\nError:\r\n' + err.message);
  }
}

// RANDOM - random integer function for selecting Game of the Week
function getRandomInt(min, max) {
      min = Math.ceil(min);
      max = Math.floor(max);
      return Math.floor(Math.random() * (max - min + 1)) + min;
}
