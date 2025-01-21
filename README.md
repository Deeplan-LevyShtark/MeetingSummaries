# MeetingSummaries

## Summary

Short summary on functionality and used technologies.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.18.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Solution

| Solution | Author(s)
| -------- | ---------
| MeetingSummaries | DeePlan

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |

## Technology versions used

* Node.js - v18.17.1
* Gulp    - 3.0.0
* Npm     - 3.0.0

## Set your environment

Please follow this guide in order to set up your SharePoint Framework development environment:
[SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-development-environment)

## Preparations

- Since we use @pnp/sp in this library, every time we add or update a package we need to follow this guide:
https://pnp.github.io/pnpjs/v2/SPFx-on-premises/
When using the Yeoman generator to create a SharePoint Framework 1.4.1 project targeting on-premises it installs TypeScript version 2.2.2 (SP2016) or 2.4.2/2.4.1 (SP2019). Unfortunately this library relies on 3.6.4 or later due to extensive use of default values for generic type parameters in the libraries. To work around this limitation you can follow the steps in this article.

## Debug

- Open the command line and navigate to folder where this readme exists
- In the command line run:
  - `npm install`
  - `gulp trust-dev-cert`
  - `npm start`

## Deploy

- Move to folder where this readme exists
- In the command line run:
  - `gulp clean`
  - `gulp bundle --ship`
  - `gulp package-solution --ship`
- Upload .sppkg file from sharepoint\solution to your tenant App Catalog
  E.g.: https://<tenant>.sharepoint.com/sites/AppCatalog/AppCatalog
- Only on the first upload: you need to approve Graph API request in the office365 admin center:
  office365 admin >>> Advanced (right pannel) >>> Api access

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development