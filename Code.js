const scriptProperties = PropertiesService.getScriptProperties();
const PRIVATE_KEY = scriptProperties
  .getProperty("private_key")
  .replace(/\\n/g, "\n");
const PROJECT_ID = scriptProperties.getProperty("project_id");
const CLIENT_EMAIL = scriptProperties.getProperty("client_email");
const API_KEY = scriptProperties.getProperty("OPENAIKEY");
const MODEL_TYPE = "gpt-4-0613";
const OPENAI_ENDPOINT = "https://api.openai.com/v1/chat/completions";

/**
 * Transcribes an audio file from a specified GCS URI using Google's Speech-to-Text API.
 *
 * @param {string} gcsUri - The GCS URI of the audio file to be transcribed.
 * @returns {Object|string} - An object containing the transcription and confidence, or
 *                            the operation name if the transcription is not immediately available.
 */
function transcribeAudioFromGCS(gcsUri) {
  var service = getService();

  // Check if the service has access to the API
  if (!service.hasAccess()) {
    Logger.log(service.getLastError());
    return;
  }

  // Define the endpoint and payload for the Speech-to-Text API request
  var apiUrl = "https://speech.googleapis.com/v1/speech:longrunningrecognize";
  var payload = {
    config: {
      language_code: "en-US",
    },
    audio: {
      uri: gcsUri,
    },
  };

  // Set the options for the API request
  var options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + service.getAccessToken(),
    },
    payload: JSON.stringify(payload),
  };

  // Make the API request and parse the response
  var response = UrlFetchApp.fetch(apiUrl, options);
  var jsonResponse = JSON.parse(response);

  // Handle the API response based on its content
  if (jsonResponse && jsonResponse.results && jsonResponse.results.length > 0) {
    var transcription = jsonResponse.results[0].alternatives[0].transcript;
    var confidence = jsonResponse.results[0].alternatives[0].confidence;

    Logger.log("Transcription: " + transcription);
    Logger.log("Confidence: " + confidence);

    return {
      transcription: transcription,
      confidence: confidence,
    };
  } else {
    Logger.log("No transcription available");
    var operationName = jsonResponse.name;
    Logger.log(operationName);
    return operationName;
  }
}

/**
 * Fetches the status of a long-running transcription operation from Google's Speech-to-Text API.
 *
 * @param {string} checkOperation - The operation name to check its status.
 * @returns {Object|null} - A JSON response if the operation is completed, or null if it's still in progress or there's an error.
 */
function fetchOperationStatus(checkOperation) {
  var service = getService();

  // Check if the service has access to the API
  if (!service.hasAccess()) {
    Logger.log(service.getLastError());
    return;
  }

  // Define the endpoint for fetching the status of the transcription operation
  var operationName = checkOperation;
  var apiUrl = "https://speech.googleapis.com/v1/operations/" + operationName;

  // Set the options for the API request
  var options = {
    method: "get",
    contentType: "application/json; charset=utf-8",
    headers: {
      Authorization: "Bearer " + service.getAccessToken(),
    },
  };

  // Make the API request and parse the response
  var response = UrlFetchApp.fetch(apiUrl, options);
  var jsonResponse = JSON.parse(response);

  // Handle the API response based on its content
  if (jsonResponse && jsonResponse.done) {
    Logger.log("Operation completed");
    Logger.log(JSON.stringify(jsonResponse, null, 2));
    return jsonResponse;
  } else {
    Logger.log("Operation still in progress or encountered an error");
    Logger.log(JSON.stringify(jsonResponse, null, 2));
    return null;
  }
}

/**
 * Lists all files in a specified Google Cloud Storage (GCS) bucket directory and writes the URIs to the active spreadsheet.
 *
 * @param {string} gcsPath - The GCS path in the format 'bucketName/directoryPath'.
 */
function listFilesInBucket(gcsPath) {
  // Check for a valid GCS path input
  if (!gcsPath) {
    Logger.log("No GCS path provided.");
    return;
  }

  // Split the GCS path into its constituent parts (bucket name and directory path)
  const parts = gcsPath.split("/");
  const bucketName = parts[0];
  const directoryPath = parts.slice(1).join("/");

  // Fetch and check access to the Google Cloud service
  const service = getService();
  if (!service.hasAccess()) {
    Logger.log(service.getLastError());
    return;
  }

  // Set the base URIs for API request and for file listing
  const baseUri = "gs://";
  const baseURL =
    "https://storage.googleapis.com/storage/v1/b/" +
    bucketName +
    "/o?prefix=" +
    encodeURIComponent(directoryPath);

  // Configure the options for the API request
  const options = {
    method: "GET",
    headers: {
      Authorization: "Bearer " + service.getAccessToken(),
      Accept: "application/json",
    },
    muteHttpExceptions: true,
  };

  let nextPageToken; // Token for paginated results
  const uris = [];

  do {
    // Append the nextPageToken to the URL if it exists
    const currentUrl =
      baseURL + (nextPageToken ? "&pageToken=" + nextPageToken : "");

    // Fetch and parse the API response
    const response = UrlFetchApp.fetch(currentUrl, options);
    const data = JSON.parse(response.getContentText());

    // Extract file information and construct URIs
    const files = data.items || [];
    files.forEach((file) => {
      const fileUri = baseUri + bucketName + "/" + file.name;
      uris.push([fileUri]);
    });

    // Update the nextPageToken for pagination
    nextPageToken = data.nextPageToken;
  } while (nextPageToken); // Continue fetching pages while there are more results

  // Write the constructed URIs to the active spreadsheet
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, uris.length, 1).setValues(uris);
}

/**
 * Converts a Google Cloud Storage URL to its corresponding GCS URI format.
 *
 * @param {string} url - The Google Cloud Storage URL to be converted.
 * @returns {string} - The converted GCS URI.
 */
function convertToGcsUri(url) {
  return "gs://" + url.replace("https://storage.googleapis.com/", "");
}

/**
 * Wrapper function to initiate listing of files in the specified GCS directory.
 */
function insertGsutilUris() {
  listFilesInBucket(
    "gcs-bucket-name/directory-path" // Replace with the actual GCS path
  );
}

/**
 * Initiates a batch transcription process for a list of Google Cloud Storage URIs.
 * The URIs are fetched from the first column of the active spreadsheet, and the
 * resulting operation names from the transcription process are written to the second column.
 */
function startBatchTranscription() {
  // Fetch active sheet and its last row number
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Retrieve the list of GCS URIs from the first column
  var gcsUris = sheet.getRange(1, 1, lastRow).getValues();

  // Loop through each GCS URI to start the transcription process
  for (var i = 0; i < gcsUris.length; i++) {
    var gcsUri = gcsUris[i][0];
    if (gcsUri) {
      var operationName = transcribeAudioFromGCS(gcsUri);
      if (operationName) {
        // Save the resulting operation name to the second column
        sheet.getRange(i + 1, 2).setValue(operationName);
      }
    }
  }
}

/**
 * Fetches the completed transcriptions based on operation names in the active spreadsheet.
 * Operation names are taken from the second column, and the retrieved transcriptions are
 * saved in the third column of the same row.
 */
function fetchCompletedTranscriptions() {
  // Fetch active sheet and its last row number
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Retrieve the list of operation names from the second column
  var operations = sheet.getRange(1, 2, lastRow).getValues();

  // Loop through each operation name to fetch the corresponding transcription
  for (var i = 0; i < operations.length; i++) {
    var operationName = operations[i][0];
    if (operationName) {
      var jsonResponse = fetchOperationStatus(operationName);
      if (jsonResponse && jsonResponse.done) {
        // Concatenate the retrieved transcripts
        var concatenatedTranscript = concatenateTranscripts(jsonResponse);
        // Save the concatenated transcript to the third column
        sheet.getRange(i + 1, 3).setValue(concatenatedTranscript);
      }
    }
  }
}

/**
 * Concatenates individual transcripts from a provided JSON response.
 *
 * @param {Object} jsonResponse - The JSON response containing transcription results.
 * @returns {string} - A concatenated string of all individual transcripts.
 */
function concatenateTranscripts(jsonResponse) {
  var transcripts = [];

  // Check for the presence of transcription results in the jsonResponse
  if (jsonResponse.response && jsonResponse.response.results) {
    jsonResponse.response.results.forEach(function (result) {
      // Ensure that a valid transcript exists within the results
      if (result.alternatives && result.alternatives[0]) {
        transcripts.push(result.alternatives[0].transcript);
      }
    });
  }
  // Join individual transcripts into a single string and return
  return transcripts.join(" ");
}

function processColumn(sourceColumn, targetColumn, processorFunction) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet
    .getRange(sourceColumn + "2:" + sourceColumn + sheet.getLastRow())
    .getValues();

  for (let i = 0; i < data.length; i++) {
    const textToProcess = data[i][0];
    if (
      textToProcess &&
      typeof textToProcess === "string" &&
      textToProcess.trim() !== ""
    ) {
      const targetCell = sheet.getRange(targetColumn + (i + 2));
      const existingContent = targetCell.getValue();

      if (
        !existingContent ||
        (typeof existingContent === "string" && existingContent.trim() === "")
      ) {
        const processedText = processorFunction(textToProcess);
        targetCell.setValue(processedText);
        Logger.log("Row " + (i + 2) + ": " + processedText);
        Logger.log(
          "Processed " + (i + 1) + " out of " + data.length + " rows."
        );
        Utilities.sleep(500);
      } else {
        Logger.log("Row " + (i + 2) + " already has content. Skipping.");
      }
    }
  }
}

function summarizeText(text) {
  const prompt =
    "Given the transcription, produce a concise summary.Transcriptions may have errors, so use context to interpret. Highlight key subjects, actions, and crucial details. Identify and convey recurring themes or sentiments. Clarify any ambiguities in your summary. Transcription: " +
    text;
  return fetchFromOpenAI(prompt, 0.3, 3050);
}

function insightsFromTranscript(text) {
  const prompt =
    "Analyze this phone call transcription and provide brief, actionable insights (under 100 words) on improvement: " +
    text;
  return fetchFromOpenAI(prompt, 0.3, 3050);
}

function rankText(text) {
  const prompt =
    "Based on the description of a customer service phone call, rank the experience from 0 (worst) to 10 (best). Provide only a whole number without explanation: " +
    text;
  return fetchFromOpenAI(prompt, 0.3, 3050);
}

function summarizeallRows() {
  processColumn("F", "G", summarizeText);
}

function rankAllRows() {
  processColumn("G", "H", rankText);
}

function provideInsightsAllRows() {
  processColumn("G", "I", insightsFromTranscript);
}

// Custom Functions for Sheets

function fetchFromOpenAI(prompt, temperature, maxTokens) {
  const requestBody = {
    model: MODEL_TYPE,
    messages: [{ role: "user", content: prompt }],
    temperature: temperature,
    max_tokens: maxTokens,
  };

  const requestOptions = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + API_KEY,
    },
    payload: JSON.stringify(requestBody),
  };

  const response = UrlFetchApp.fetch(OPENAI_ENDPOINT, requestOptions);
  return JSON.parse(response.getContentText())["choices"][0]["message"][
    "content"
  ];
}

// function gptProcess(text, prompt) {
//     const temperature = 0.3;
//     const maxTokens = 3000;

//     const requestBody = {
//       model: "gpt-3.5-turbo-16k",
//       messages: [{ role: "user", content: prompt }],
//       temperature,
//       max_tokens: maxTokens,
//     };

//     const requestOptions = {
//       method: "POST",
//       headers: {
//         "Content-Type": "application/json",
//         Authorization: "Bearer " + API_KEY,
//       },
//       payload: JSON.stringify(requestBody),
//     };

//     const response = UrlFetchApp.fetch(
//       "https://api.openai.com/v1/chat/completions",
//       requestOptions
//     );
//     const json = JSON.parse(response.getContentText());
//     return json["choices"][0]["message"]["content"];
//   }
