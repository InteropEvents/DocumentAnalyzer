# DevDays Asia 2022 - Use Azure Cognitive Service for Language with the OpenXML SDK

## Contents

- [Use Azure Cognitive Service for Language with the OpenXML SDK](#use-azure-cognitive-service-for-language-with-the-openxml-sdk)
- [Prerequisites](#prerequisites)
- [Get Started](#get-started)
- [Follow the Tutorial](#follow-the-tutorial)
- [Working Sample](#working-sample)
- [Next steps](#next-steps)
  - [Learn more about the OpenXML SDK](#learn-more-about-the-openxml-sdk)
  - [Learn more about Azure Cognitive Services for Language](#learn-more-about-azure-cognitive-services-for-language)

## Use Azure Cognitive Service for Language with the OpenXML SDK

In this repo there is a sample console app that uses the [OpenXML SDK](https://github.com/OfficeDev/Open-XML-SDK) and [Azure Cognitive Service for Language](https://docs.microsoft.com/en-us/azure/cognitive-services/language-service/overview) with tutorial modules showing how to write the code to open a Word document, examine its contents for PII and save a copy with the PII redacted.

## Prerequisites

- [Visual Studio](https://visualstudio.microsoft.com/downloads/) (community edition is OK)
- [.NET 6.0](https://dotnet.microsoft.com/en-us/download)
- [git command line tools](https://git-scm.com/downloads)
- Word Processing app that can open .docx files such as [LibreOffice](https://www.libreoffice.org/download/download-libreoffice/) or [Word](https://www.microsoft.com/en-us/microsoft-365/word)

## Get started

- Clone this repository to your local system.

  `git clone xxxxxxxxxxxxxxxxxxxxxxxx`

    **Pro tip:** Clone the repo low in your folder hierarchy to avoid path length issues e.g. `C:\myrepos`

- Create an Azure account and Language Services resource and copy the API key and endpoint by following the steps in the [setup document](./docs/setup.md).

  *If you already have a API key and endpoint provided to you, skip this step.*

## Follow the tutorial

Now that you have cloned the repo and have your API key and endpoint, you're ready to create your application.

1. First follow [this document](./docs/create-project.md) to create your console application with Visual Studio.

2. Next install the the dependencies with [this document](./docs/install-packages.md).

4. Then follow [this document](./docs/write-code.md) to write the application code.

## Working sample

For a working example of this app:

- Clone this repo

- Open `DocumentAnalyzer.csproj` with Visual Studio

- Replace the placeholders for the API key, endpoint, and file path in `Program.cs`

- Press `F5`

## Next steps

### Learn more about the OpenXML SDK

- [Tutorials for the Open XML SDK](https://docs.microsoft.com/en-us/office/open-xml/how-do-i) are available on docs.microsoft.

- [Example apps](https://github.com/OfficeDev/Open-XML-SDK/tree/main/samples) are also available.

- Contribute to the project on the [OpenXML SDK repo](https://github.com/OfficeDev/Open-XML-SDK/)

### Learn more about Azure Cognitive Services for Language

- The [documentation for Azure Cognitive Services for Language](https://docs.microsoft.com/en-us/azure/cognitive-services/language-service/overview) is available at docs.microsoft.

- [Sample code](https://docs.microsoft.com/en-us/azure/cognitive-services/language-service/concepts/developer-guide?tabs=language-studio) for C#, Java, JavaScript, and Python are also available at docs.microsoft.
