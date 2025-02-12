# Google Sheets to PowerPoint Converter

![License](https://img.shields.io/badge/License-MIT-blue.svg)
![Python Version](https://img.shields.io/badge/Python-3.7%2B-blue.svg)
![GitHub stars](https://img.shields.io/github/stars/yourusername/docx-to-google-sheets-converter?style=social)

## Overview

The **Google Sheets to PowerPoint Converter** is a versatile tool designed to transform data stored in Google Sheets into polished PowerPoint presentations. This application is particularly useful for professionals and educators who need to quickly create presentations based on dynamic data updated in real-time on Google Sheets.

## Features

- **Automated Conversion:** Seamlessly fetches data from Google Sheets and generates comprehensive PowerPoint presentations.
- **Customizable Templates:** Supports custom PowerPoint templates to maintain consistency in presentation style and branding.
- **Real-Time Updates:** Capable of updating presentations as the source Google Sheet changes, keeping the presentation up-to-date.
- **Interactive GUI:** User-friendly graphical interface that makes it simple to link Google Sheets, choose templates, and initiate conversions.
- **Scalable:** Efficiently processes large datasets without performance bottlenecks.
- **Open Source:** Available for community use and contributions, fostering enhancement and diversification of features.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
  - [Install Python Dependencies](#install-python-dependencies)
- [Setting Up Google API Credentials](#setting-up-google-api-credentials)
  - [Create a Project in Google Cloud Console](#create-a-project-in-google-cloud-console)
  - [Enable Google Sheets API](#enable-google-sheets-api)
  - [Create OAuth 2.0 Credentials](#create-oauth-20-credentials)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)

## Prerequisites

Before you begin, ensure you have the following installed on your system:

- **Python 3.7+**  
  Download and install Python from the [official website](https://www.python.org/downloads/).

- **Pandoc**  
  Pandoc is required for converting DOCX files to LaTeX.
  - **Windows:** Download from [Pandoc's official website](https://pandoc.org/installing.html).
  - **macOS:** Install using Homebrew:
    ```bash
    brew install pandoc
    ```
  - **Linux:** Install via your distribution's package manager, e.g., for Debian/Ubuntu:
    ```bash
    sudo apt-get install pandoc
    ```

- **LaTeX Distribution**  
  Required for rendering LaTeX equations.
  - **Windows:** Install [MiKTeX](https://miktex.org/download).
  - **macOS:** Install [MacTeX](http://www.tug.org/mactex/).
  - **Linux:** Install TeX Live via your distribution's package manager, e.g., for Debian/Ubuntu:
    ```bash
    sudo apt-get install texlive-full
    ```

## Installation

### Install Python Dependencies

Creating a virtual environment helps manage dependencies and avoid conflicts with other projects.

```bash
pip install google-auth google-auth-oauthlib google-api-python-client python-pptx matplotlib Pillow
```

## Setting Up Google API Credentials

### Create a Project in Google Cloud Console
- Navigate to the [Google Cloud Console](https://console.cloud.google.com/)
- Click on the project dropdown and select New Project.
- Enter a project name and click Create.
### Enable Google Sheets API
- Within your project dashboard, go to APIs & Services > Library.
- Search for "Google Sheets API" and click on it.
- Click Enable.

### Create OAuth 2.0 Credentials

-  Go to APIs & Services > Credentials.
-  Click on + CREATE CREDENTIALS and select OAuth client ID.
-  If prompted, configure the OAuth consent screen by providing necessary information.
-  Choose Desktop app as the application type and provide a name.
-  Click Create.
-  Download the JSON file, rename it to credentials.json, and place it in the root   directory of your project.

## Contributing

Contributions make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are greatly appreciated.

### Fork the Project

- Create Your Feature Branch

```bash
git checkout -b feature/AmazingFeature
```

- Commit Your Changes

```bash
git commit -m 'Add some AmazingFeature'
```
- Push to the Branch

```bash
git push origin feature/AmazingFeature
```

- Open a Pull Request

## License
Distributed under the MIT License. See LICENSE for more information.

## Contact
Rifat Ahmed - rifu.cse.bubt@gmail.com

Project Link: https://github.com/ress-rifu/Gsheet2PPTX/
