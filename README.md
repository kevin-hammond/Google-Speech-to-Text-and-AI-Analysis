# Google Cloud Speech-to-Text and OpenAI Integration

This script is designed to transcribe audio files in a Google Cloud Storage (GCS) bucket, analyze the transcriptions, and write the results to a Google Sheets spreadsheet. It uses Google's Speech-to-Text API for transcription and OpenAI's API for analysis.

Note: the script is designed to be run in Google Apps Script and the file extension should be `.gs` not `.js`. Change this in the script editor if necessary.

## Prerequisites

1. A Google Cloud project with the Speech-to-Text API enabled.
2. A Google Cloud Storage bucket with audio files.
3. A Google Sheets spreadsheet to store the results.
4. OpenAI API key.

## Configuration

1. Set up the required properties in the script's properties service:
   - `private_key`: The private key for the Google Cloud project.
   - `project_id`: The project ID for the Google Cloud project.
   - `client_email`: The client email for the Google Cloud project.
   - `OPENAIKEY`: The OpenAI API key.

## Functions

- `transcribeAudioFromGCS(gcsUri)`: Transcribes an audio file from a specified GCS URI using Google's Speech-to-Text API.
- `fetchOperationStatus(checkOperation)`: Fetches the status of a long-running transcription operation from Google's Speech-to-Text API.
- `listFilesInBucket(gcsPath)`: Lists all files in a specified GCS bucket directory and writes the URIs to the active spreadsheet.
- `convertToGcsUri(url)`: Converts a Google Cloud Storage URL to its corresponding GCS URI format.
- `insertGsutilUris()`: Wrapper function to initiate listing of files in the specified GCS directory.
- `startBatchTranscription()`: Initiates a batch transcription process for a list of Google Cloud Storage URIs.
- `fetchCompletedTranscriptions()`: Fetches the completed transcriptions based on operation names in the active spreadsheet.
- `concatenateTranscripts(jsonResponse)`: Concatenates individual transcripts from a provided JSON response.
- `summarizeText(text)`: Summarizes a given text using OpenAI's API.
- `insightsFromTranscript(text)`: Provides insights from a given transcript using OpenAI's API.
- `rankText(text)`: Ranks a given text describing a customer service phone call using OpenAI's API.
- `summarizeallRows()`: Summarizes all text in a specified column.
- `rankAllRows()`: Ranks all text in a specified column.
- `provideInsightsAllRows()`: Provides insights for all text in a specified column.
- `fetchFromOpenAI(prompt, temperature, maxTokens)`: Fetches a response from OpenAI's API given a prompt, temperature, and maximum number of tokens.

## Usage

1. Run `insertGsutilUris()` to list all files in the specified GCS directory.
2. Run `startBatchTranscription()` to initiate a batch transcription process for the listed files.
3. Run `fetchCompletedTranscriptions()` to fetch the completed transcriptions.
4. Run `summarizeallRows()`, `rankAllRows()`, or `provideInsightsAllRows()` to analyze the transcriptions.

For custom functions in Google Sheets:

- `summarizeText(text)`: Summarizes a given text.
- `insightsFromTranscript(text)`: Provides insights from a given transcript.
- `rankText(text)`: Ranks a given text describing a customer service phone call.

Use these functions in a cell by entering `=summarizeText(A1)`, `=insightsFromTranscript(A1)`, or `=rankText(A1)`, replacing `A1` with the cell reference containing the text to analyze.
