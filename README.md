# React Page Creator

## Summary
This web part enables creating new Site Pages or News Posts across different site collections in a SharePoint tenant. It also gives the option to create the new pages as blank pages or using any of the template pages.

![react-command-get-thumbnail](./assets/GetThumbnail.gif)

## Used SharePoint Framework Version

![1.9.1](https://img.shields.io/badge/version-1.9.1-green.svg)

## Applies to

* [SharePoint Framework](https://dev.office.com/sharepoint)
* [Office 365 tenant](https://dev.office.com/sharepoint/docs/spfx/set-up-your-development-environment)

## Prerequisites

There are no pre-requisites.

## Solution

Solution|Author(s)
--------|---------
react-page-creator | [Aakash Bhardwaj](https://twitter.com/aakash_316)

## Version history

Version|Date|Comments
-------|----|--------
1.0| January 1, 2020|Initial release

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

* Clone this repository
* `npm install`
* `gulp bundle --ship`
* `gulp package-solution --ship`
* Add to Site Collection App Catalog and Install the App
* Go to the API Management section in the new SharePoint Admin Center (https://{tenantname}-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement)
* Approve the permission request for Sites.Read.All to Microsoft Graph

## Features

This web part has the following features:

* Creating new Site Pages and News Posts across different site collections
* Getting available template pages in the selected site collection
* Getting sites being followed by the current user using Graph API
