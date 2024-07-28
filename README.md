# Jira Issues Report Generator

This project generates an Excel report of Jira issues for specified team members within a given month. The report includes various details about each issue and applies conditional formatting to certain cells.

## Prerequisites

- Node.js
- npm
- A Jira access token is required to access the Jira API.

## Setup

1. Clone the repository:

   ```sh
   git clone https://github.com/son-quach-topicus/jira-issues-report-generator.git
   cd jira-issues-report-generator
   ```

2. Install the dependencies:

   ```sh
   npm install
   ```

3. Create a `.env` file in the root directory and add the following environment variables:

   ```env
   JIRA_USERNAME=your-email@example.com
   JIRA_ACCESS_TOKEN=your-jira-access-token
   JIRA_TEAM_NAME=your-team-name
   JIRA_TEAM_MEMBERS=member1,member2
   JIRA_PROJECT_NAME=your-project-name
   JIRA_STATUS=Done,Closed,Scheduled
   JIRA_URL=https://jira.topicus.nl
   JIRA_STATUS_CATEGORY=Done
   JIRA_ORDER_BY=updated
   JIRA_ORDER_DIRECTION=desc
   ```

## Usage

To generate the report, run the following command:

```sh
npm run export
```
