# CES

The Visual Basic for Applications (VBA) code for the Course Evaluation Surveys.

The code processes the student demographic/registration data to format and prepare it for upload to the survey software.

The data output from Qualtrics survey software is saved and processed, as follows:

*Pre-Processing 
  * Ensuring all reportable criteria are matched to the data
  * Splitting sheets into reportable courses and modules

*Course and Module Reporting
 * Generating a report on responses for each individual study year for each course (192 courses)
 * Generating a report on responses for each module (732 modules)
 * Ensuring that the cohort size threshold had been met to ensure anonymity
 * Saving these reports in folders by School to be circulated

*Summary Reports
 * Generating summary reports by Department and School for both course and module responses

*Checksheets
 * Generating checking sheets for administrative staff to vet the free-text comments supplied prior to circulation
