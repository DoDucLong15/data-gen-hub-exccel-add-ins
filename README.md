# Data Gen Hub Excel Add-in

## Overview

This project is an Office Add-in for Microsoft Excel that allows users to map Excel sheet data to database fields, configure multiple sheets, and generate a specification file (`spec.json`) for further data processing or integration. The add-in provides an interactive task pane UI for defining mappings, configuring sheet properties, and exporting the configuration as a JSON file.

## Features
- Map Excel cells to database fields with a flexible UI.
- Configure multiple sheets, each with its own mapping and settings.
- Support for predefined database tables and fields.
- Export all configuration as a JSON specification file (`spec.json`).
- Built using TypeScript, Office.js, Webpack, and modern front-end tooling.

## Prerequisites
- [Node.js](https://nodejs.org/) (version 14 or higher recommended)
- [npm](https://www.npmjs.com/) (comes with Node.js)
- Microsoft Excel (desktop or web version with Office Add-ins support)

## Getting Started

### 1. Clone the Repository
```
git clone <your-repository-url>
cd data-gen-hub/add-ins/data-gen-hub-excel
```

### 2. Install Dependencies
```
npm install
```

### 3. Build the Add-in
- For production build:
  ```
  npm run build
  ```
- For development build (with hot reload):
  ```
  npm run dev-server
  ```

### 4. Sideload the Add-in into Excel
1. Ensure you have Excel installed (desktop or web version).
2. Run the following command to start the local server and sideload the add-in:
   ```
   npm start
   ```
   This will launch Excel and load the add-in for testing.

3. If prompted, trust the self-signed certificate generated for local development.

### 5. Using the Add-in
- Open the task pane from the Excel ribbon.
- Use the UI to add sheets, map Excel cells to database fields, and configure settings.
- Click the export button to generate and download the `spec.json` file with your configuration.

## Development Scripts
- `npm run build` — Build the add-in for production.
- `npm run build:dev` — Build the add-in for development.
- `npm run dev-server` — Start the development server with hot reload.
- `npm start` — Sideload and debug the add-in in Excel.
- `npm run stop` — Stop the debugging session.
- `npm run lint` — Lint the codebase.
- `npm run lint:fix` — Fix lint errors automatically.

## Project Structure
- `src/taskpane/` — Main task pane UI and logic.
- `src/commands/` — Command interface for the add-in.
- `manifest.xml` — Office Add-in manifest file.
- `webpack.config.js` — Webpack configuration.

## License
This project is licensed under the MIT License.

---

If you have any questions or need further assistance, please refer to the source code or open an issue in your repository.
