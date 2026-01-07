# Field Audit Trail PCF Control

A modern Power Apps Component Framework (PCF) control that provides instant visibility into a field's audit history directly within the form interface.

## 🚀 Features

- **Inline History Access**: View field changes without leaving the record or opening the global Audit History.
- **Detailed Audit Info**: Displays the last 5 changes including User, Timestamp, Operation (Create/Update), and the specific Old/New values.
- **Smart Parsing**: Automatically handles both JSON and XML `changedata` formats from Dataverse.
- **Metadata Aware**: Resolves OptionSet and TwoOptions values to their display labels for better readability.
- **Premium UI**: Built with Fluent UI, featuring a sleek, blurred callout (glassmorphism) and smooth transitions.
- **Interactive**: Hover effects and easy-to-read activity items.

## 🛠️ Technology Stack

- **Framework**: Power Apps Component Framework (PCF)
- **Library**: React
- **Styling**: Fluent UI (@fluentui/react)
- **Language**: TypeScript

## 📦 Prerequisites

- [Node.js](https://nodejs.org/) (Recommended version: 16.x or 18.x)
- [Microsoft Power Platform CLI](https://learn.microsoft.com/en-us/power-platform/developer/cli/introduction)

## 🏗️ Getting Started

### Installation

1. Clone the repository or download the source code.
2. Install dependencies:
   ```bash
   npm install
   ```

### Development & Testing

Run the control in the local test harness:
```bash
npm start
```

### Building the Control

To build the control for production:
```bash
npm run build
```

## 🛠️ Configuration

### Manifest Properties
- **Field Value**: The data-bound field that this control will track and display.

### Features Used
- **WebAPI**: Required to fetch audit records from the `audit` entity.
- **Utility**: Used for UI utilities.

## 📝 How it Works

1. The control renders as a `TextField` paired with a history icon.
2. When the icon is clicked, it fetches the `audit` records for the current record ID using the Dataverse Web API.
3. It filters the audit records to find entries where the specific field was changed.
4. It parses the `changedata` (supporting both older XML and newer JSON formats).
5. It cross-references OptionSet metadata to ensure human-readable labels are shown instead of integer values.
6. The results are displayed in a premium Fluent UI `Callout`.

## 📜 License

This project is licensed under the MIT License.

## 📞 Contact

- **Author**: Advic Tech
- **Email**: developer@advic.io
- **Issues**: Report bugs or suggestions on the GitHub Issues page.

