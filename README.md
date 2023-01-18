# Hatch Test Add in

## How to install and set up add in
1. Install [node.js](https://nodejs.org/en/) and [git](https://git-scm.com/downloads)
2. Clone the respository
3. Run `npm run build` in terminal
4. Run `npm run start:desktop`
5. Open Excel workbook, navigate to **Insert**, select **My Add-ins**, and select the **Hatch test addin** under **Developer Add-ins**

## How to add debugger
1. Go to **Home** and select **Show Taskpane**
2. Select the **<** in the **Hatch test addin** taskpane and select **Attach Debugger**
3. Nagivate to the console

## How to setup the workbook to use getStreamValue function
1. The sheet containing the stream values to be searched should be named **Test**
2. A name should be created for the search range including headers using the **Name Manager**. This range should be named **PropTypeValue**. 
3. Alternatively, the sheet name and search range name can be modified in the `getRange` function (lines 38 and 39 of the `functions.js` file)
4. The headers need to be in the first row of the search range. 
