# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Chrome browser extension (Manifest V3) designed to extract table data from the EbizWare WMS system at `https://ebizware.yunwms.com/` and export it to CSV format. The extension targets a specific logistics management system and provides data extraction functionality for warehouse management records.

## Architecture

The extension follows the standard Chrome extension architecture with three main components:

- **manifest.json**: Extension configuration defining permissions, content script targets, and service worker
- **background.js**: Service worker that handles extension icon clicks and injects the content script
- **content.js**: Content script that runs on the target site, performs DOM manipulation, data extraction, and CSV generation

## Key Functionality

The extension specifically targets table data extraction from EbizWare WMS forms with these data fields:
- 认领单号 (Claim ID)
- 仓库 (Warehouse) 
- 跟踪号 (Tracking Number)
- 包裹数量 (Package Quantity)
- 参考号 (Reference Number)
- 有效时间 (Valid Time)
- 认领时间 (Claim Time)
- 创建时间 (Created Time)
- 完成时间 (Completed Time)
- 更新时间 (Updated Time)
- 备注 (Notes)

## Development Commands

This is a browser extension project with no build system or package.json. Development involves:

1. **Loading the extension**: Load the folder directly in Chrome's developer mode at `chrome://extensions/`
2. **Testing**: Navigate to `https://ebizware.yunwms.com/` and click the extension icon to test functionality
3. **Debugging**: Use Chrome DevTools console for content script debugging and extension's service worker inspector for background script debugging

## Target Site Integration

The extension is configured to run only on `https://ebizware.yunwms.com/` and expects:
- A form with id `listForm` containing a table
- Table rows with specific cell structure and HTML patterns for data extraction
- Chinese language interface elements

## Content Script Behavior

The content script:
- Creates a floating "导出CSV" (Export CSV) button in the top-right corner
- Prevents duplicate button creation through ID checking
- Extracts data using querySelector and regex pattern matching
- Generates CSV with UTF-8 BOM encoding
- Provides user feedback through browser alerts