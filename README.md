# VBA Automation Scripts for Microsoft Planner and OpenAI GPT Integration

This repository contains two VBA scripts that automate tasks using Microsoft Planner and OpenAI GPT-3.5. These scripts help streamline task creation in Microsoft Planner from an Excel spreadsheet and generate text completions using OpenAI's API.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Scripts](#scripts)
  - [Microsoft Planner Task Automation](#microsoft-planner-task-automation)
  - [OpenAI Text Completion](#openai-text-completion)
- [Usage](#usage)
- [License](#license)

## Overview

### Microsoft Planner Task Automation
This script automates the creation of tasks in Microsoft Planner based on data from an Excel spreadsheet. It uses Microsoft Graph API to interact with Microsoft Planner and supports setting task titles, due dates, assignees, and reminders.

### OpenAI Text Completion
This script integrates with OpenAI's GPT-3.5 API to generate text completions based on selected rows in an Excel spreadsheet. The results are then displayed in a specified column.

## Prerequisites

- **Excel**: Both scripts are designed to run in Microsoft Excel.
- **VBA JSON Parser**: The scripts use a JSON parser (`JsonConverter.bas`). Make sure this module is included in your VBA project. You can find it in this repository or download it from [VBA-JSON GitHub](https://github.com/VBA-tools/VBA-JSON).
- **Microsoft Graph API Access**: For the Planner script, you'll need access to Microsoft Graph API with appropriate permissions.
- **OpenAI API Key**: For the OpenAI script, you'll need an API key from OpenAI. You can sign up for a key at [OpenAI](https://openai.com/api/).

## Installation

1. **Clone the Repository**: Clone this repository to your local machine or download the ZIP file and extract it.

2. **Add the JSON Parser**: Import the `JsonConverter.bas` file into your VBA project.

3. **Set Up API Credentials**:
   - For Microsoft Planner: Update the `CLIENT_ID`, `CLIENT_SECRET`, and `TENANT_ID` constants in the Planner script.
   - For OpenAI: Add your OpenAI API key to the `API_KEY` constant in the OpenAI script.

## Scripts

### Microsoft Planner Task Automation

This script automates the process of creating tasks in Microsoft Planner based on data in an Excel spreadsheet. It:
- Retrieves an access token using OAuth2.
- Creates tasks in a specified Planner plan and bucket.
- Supports assigning tasks to users based on their email.
- Creates reminder tasks 4 weeks before the due date.

#### Key Constants
- `CLIENT_ID`, `CLIENT_SECRET`, `TENANT_ID`: Credentials for accessing the Microsoft Graph API.
- `PLAN_ID`, `BUCKET_ID`: Identifiers for the target plan and bucket in Microsoft Planner.
- `DEFAULT_ASSIGNEE_ID`: A fallback user ID in case the assignee's email is not found.

### OpenAI Text Completion

This script integrates with OpenAI's GPT-3.5 model to generate text completions. It reads data from the Excel spreadsheet, sends it to OpenAI, and displays the generated text in the result column.

#### Key Constants
- `API_ENDPOINT`: The endpoint for the OpenAI completions API.
- `MODEL`: The model to use (e.g., `gpt-3.5-turbo-instruct`).
- `MAX_TOKENS`, `TEMPERATURE`: Parameters to control the behavior of the OpenAI model.

## Usage

1. **Microsoft Planner Task Automation**:
   - Select the rows in your Excel sheet that contain the task data.
   - Run the `CreatePlannerTasks` macro.
   - Check the status of task creation in the status column of your sheet.

2. **OpenAI Text Completion**:
   - Select the rows containing the input and prompt data.
   - Run the `OpenAI_Completion` macro.
   - The generated text will be displayed in the result column.

## License
MIT LICENSE
