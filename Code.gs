function onOpen() {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu('Shifter')
    .addItem('Validate', 'validateParams')
    .addItem('Solve', 'runSolver')
    .addToUi()
}

function mapHeaderToIndex(row) {
  const paramHeaderToIndex = {}
  for (let i = 0; i < row.length; ++i) {
    paramHeaderToIndex[row[i]] = i
  }
  return paramHeaderToIndex
}

function validateParams() {
  const allParams = ['Shift', 'Category', 'Minimum assigned', 'Maximum assigned', 'Hours', 'Write results to']
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const config = spreadsheet.getSheetByName('Shifter')
    if (config === null) throw new Error('You are missing the "Shifter" sheet')
    
    const paramsRange = spreadsheet.getRangeByName('Shifts')
    if (paramsRange === null) throw new Error('You are missing the "Shifts" named range, which lists the shifts and their constraints')
    
    const constraintRange = spreadsheet.getRangeByName('AvailabilityResponses')
    if (constraintRange === null) throw new Error('You are missing the "AvailabilityResponses" named range, which contains the name of the sheet with form responses')

    const preferredRange = spreadsheet.getRangeByName('PreferredGreen')
    if (preferredRange === null) throw new Error('You are missing the "PreferredGreen" named range, which specifies whether assignments following preferences are colored green')

    const minHours = spreadsheet.getRangeByName('MinHours')
    if (minHours === null) throw new Error('You are missing the "MinHours" named range, which specifies the minimum total number of hours assigned to each worker')
    if (typeof minHours.getValue() !== 'number' || minHours.getValue() < 0) {
      minHours.activate();
      throw new Error('The MinHours parameter is negative or not a number')
    }
    const maxHours = spreadsheet.getRangeByName('MaxHours')
    if (maxHours === null) throw new Error('You are missing the "MaxHours" named range, which specifies the maximum total number of hours assigned to each worker')
    if (typeof maxHours.getValue() !== 'number' || maxHours.getValue() < 0) {
      maxHours.activate();
      throw new Error('The MaxHours parameter is negative or not a number')
    }
    if (maxHours.getValue() < minHours.getValue()) {
      throw new Error('The MaxHours parameter is less than the MinHours parameter')
    }
    
    const responses = spreadsheet.getSheetByName(constraintRange.getDisplayValue())
    if (responses === null) throw new Error(`AvailabilityResponses does not name an existing sheet in this spreadsheet. The sheet ${constraintRange.getDisplayValue()} does not exist`)

    const idColumnRange = spreadsheet.getRangeByName('IdColumn')
    if (idColumnRange === null)
      throw new Error('You are missing the "IdColumn" named range, which specifies the header of the column containing unique identifiers for each person in ' + constraintRange.getDisplayValue())
    
    const displayColumnRange = spreadsheet.getRangeByName('DisplayColumn')
    if (displayColumnRange === null)
      throw new Error('You are missing the "DisplayColumn" named range, which specifies the header of the column containing names for each person in ' + constraintRange.getDisplayValue() + ' to be displayed in the output schedule')

    const params = paramsRange.getDisplayValues()
    const missingParams = allParams.filter(it => !params[0].includes(it))
    if (missingParams.length > 0) throw new Error(`You are missing some parameters in Shifts: ${missingParams.join(", ")}`)
    
    const paramHeaderToIndex = mapHeaderToIndex(params[0])

    const expectedShifts = []
    for (let row = 1; params[row][0] != ''; ++row) {
      expectedShifts.push(params[row][paramHeaderToIndex['Shift']])
    }

    const responseData = responses.getDataRange().getDisplayValues()
    const missingShiftData = expectedShifts.filter(shift => !responseData[0].includes(shift))
    if (missingShiftData.length > 0) throw new Error(`You are missing response columns in ${constraintRange.getDisplayValue()} for the specified shifts: ${missingShiftData.join(", ")}`)
    const responseDataHeaders = mapHeaderToIndex(responseData[0])

    // Check parameters
    for (let row = 1; params[row][0] != ''; ++row) {
      const minVal = params[row][paramHeaderToIndex['Minimum assigned']]
      const maxVal = params[row][paramHeaderToIndex['Maximum assigned']]
      const hours = params[row][paramHeaderToIndex['Hours']]
      const assignTo = params[row][paramHeaderToIndex['Write results to']]
      const allDigits = /^\d+$/
      if (minVal.match(allDigits) === null) {
        paramsRange.getCell(row + 1, paramHeaderToIndex['Minimum assigned'] + 1).activate()
        throw new Error(`The minimum assigned value for ${params[row][paramHeaderToIndex['Shift']]} must be an integer`)
      }
      if (maxVal.match(allDigits) === null) {
        paramsRange.getCell(row + 1, paramHeaderToIndex['Maximum assigned'] + 1).activate()
        throw new Error(`The maximum assigned value for ${params[row][paramHeaderToIndex['Shift']]} must be an integer`)
      }
      const hoursFloat = parseFloat(hours)
      if (isNaN(hoursFloat) || hoursFloat <= 0) {
        paramsRange.getCell(row + 1, paramHeaderToIndex["Hours"] + 1).activate()
        throw new Error(`The number of hours for ${params[row][paramHeaderToIndex['Shift']]} must be a positive real number`)
      }
      if (parseInt(minVal) > parseInt(maxVal)) {
        paramsRange.getCell(row + 1, paramHeaderToIndex['Shift'] + 1).activate()
        throw new Error(`The selected cell's shift has its Minimum assigned (${minVal}) higher than Maximum assigned (${maxVal})`)
      }
      const a1notation = /^[A-Z]+\d+$/
      if (assignTo.match(a1notation) === null) {
        paramsRange.getCell(row + 1, paramHeaderToIndex['Write results to'] + 1).activate()
        throw new Error(`The selected cell is not in proper A1 notation`)
      }
    }
    
    if (!responseData[0].includes(idColumnRange.getDisplayValue()))
      throw new Error(`You are missing the "${idColumnRange.getDisplayValue()}" response column, which you have specified as the column with each person's unique ID`)
    if (!responseData[0].includes(displayColumnRange.getDisplayValue()))
      throw new Error(`You are missing the "${displayColumnRange.getDisplayValue()}" response column, which you have specified as the column with each person's display name in the schedule`)
    if (new Set(responseData[0]).size !== responseData[0].length)
      throw new Error(`The column headers of the responses sheet are not unique`)
    
    const resIdColumnIndex = responseDataHeaders[idColumnRange.getDisplayValue()]
    const resDisplayColumnIndex = responseDataHeaders[displayColumnRange.getDisplayValue()]
    const validAvailability = new Set(['Unavailable', 'Available', 'Preferred'])
    const idSet = new Set()
    for (let row = 1; row < responseData.length && responseData[row].join('').trim().length > 0; ++row) {
      const rowId = responseData[row][resIdColumnIndex].toString()
      if (rowId.trim().length === 0) {
        responses.getRange(row + 1, resIdColumnIndex + 1).activate()
        throw new Error(`The selected worker ID is missing`)
      }
      if (responseData[row][resDisplayColumnIndex].toString().trim().length === 0) {
        responses.getRange(row + 1, resDisplayColumnIndex + 1).activate()
        throw new Error(`The selected worker display name value is missing`)
      }
      let availableSlots = 0
      for (const shift of expectedShifts) {
        const column = responseDataHeaders[shift]
        const availability = responseData[row][column]
        if (!validAvailability.has(availability)) {
          responses.getRange(row + 1, column + 1).activate()
          throw new Error(`The selected worker availability value is invalid. Valid values are: ${Array.from(validAvailability).join(', ')}`)
        }
        if (availability !== 'Unavailable') ++availableSlots
      }
      if (availableSlots === 0) {
        responses.getRange(row + 1, resIdColumnIndex + 1).activate()
        throw new Error(`${responseData[row][resDisplayColumnIndex]} (${responseData[row][resIdColumnIndex]}) is not available for any shift`)
      }
      if (idSet.has(rowId)) {
        responses.getRange(row + 1, resIdColumnIndex + 1).activate()
        throw new Error(`The ID "${rowId}" is duplicated in the responses`)
      }
      idSet.add(rowId)
    }
    

    SpreadsheetApp.getUi().alert('Shifter says...', `Looks good! 👍`, SpreadsheetApp.getUi().ButtonSet.OK)
  } catch (e) {
    SpreadsheetApp.getUi().alert('Shifter says...', `${e}\n\n${e?.stack}`, SpreadsheetApp.getUi().ButtonSet.OK)
  }
}

function collectParams() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const outputSheet = spreadsheet.getSheetByName('Shifter')
  const paramsRange = spreadsheet.getRangeByName('Shifts')
  const paramsData = paramsRange.getDisplayValues()
  const paramHeaderToIndex = mapHeaderToIndex(paramsData[0])
  const params = {}
  for (let row = 1; paramsData[row][0] != ''; ++row) {
    const param = {
      name: paramsData[row][paramHeaderToIndex['Shift']],
      category: paramsData[row][paramHeaderToIndex['Category']],
      min: parseInt(paramsData[row][paramHeaderToIndex['Minimum assigned']]),
      max: parseInt(paramsData[row][paramHeaderToIndex['Maximum assigned']]),
      hours: parseFloat(paramsData[row][paramHeaderToIndex['Hours']]),
      outputRange: outputSheet.getRange(paramsData[row][paramHeaderToIndex['Write results to']])
    }
    params[param.name] = param
  }
  return params
}

/**
 * @return {string[][]}
 */
function collectResponseArray() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const constraintRange = spreadsheet.getRangeByName('AvailabilityResponses')
    const responses = spreadsheet.getSheetByName(constraintRange.getDisplayValue())
    return responses.getDataRange().getDisplayValues()
}

function getVariableName(person, shift) {
  return `${person}_${shift}`
}

/**
 * @return {{idColumn: string, displayColumn: string, preferredGreen: boolean, minHours: number, maxHours: number}}
 */
function collectConfig() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const idColumn = spreadsheet.getRangeByName('IdColumn').getDisplayValue()
  const displayColumn = spreadsheet.getRangeByName('DisplayColumn').getDisplayValue()
  const preferredGreenStr = spreadsheet.getRangeByName('PreferredGreen').getDisplayValue()
  const preferredGreen = !['0', '', 'FALSE'].includes(preferredGreenStr.toString().toUpperCase())
  const minHours = spreadsheet.getRangeByName('MinHours').getValue()
  const maxHours = spreadsheet.getRangeByName('MaxHours').getValue()
  return {idColumn, displayColumn, preferredGreen, minHours, maxHours}
}

/**
 * @param opt {LinearOptimizationService.LinearOptimizationEngine}
 * @param responses {string[][]}
 * @param headers {{[key: string]: number}}
 * @param params {{[name: string]: any}}
 * @param config {{idColumn: string}}
 */
function makeAvailableVariables(opt, responses, headers, params, config) {
  for (let row = 1; row < responses.length; ++row) {
    for (const paramName of Object.keys(params)) {
      const response = responses[row][headers[paramName]]
      const varName = getVariableName(responses[row][headers[config.idColumn]], paramName)
      if (response !== 'Unavailable') {
        opt.addVariable(varName, 0, 1, LinearOptimizationService.VariableType.INTEGER)
      }
      if (response === 'Preferred') {
        opt.setObjectiveCoefficient(varName, 1)
      }
    }
  }
}

/**
 * @param opt {LinearOptimizationService.LinearOptimizationEngine}
 * @param responses {string[][]}
 * @param headers {{[key: string]: number}}
 * @param params {{[name: string]: any}}
 * @param config {{idColumn: string, minHours: number, maxHours: number}}
 */
function constrainHoursPerPerson(opt, responses, headers, params, config) {
  for (let row = 1; row < responses.length; ++row) {
    const shiftConstraint = opt.addConstraint(config.minHours, config.maxHours)
    for (const paramName of Object.keys(params)) {
      const response = responses[row][headers[paramName]]
      const varName = getVariableName(responses[row][headers[config.idColumn]], paramName)
      if (response !== 'Unavailable') {
        shiftConstraint.setCoefficient(varName, params[paramName].hours)
      }
    }
  }
}

/**
 * @param opt {LinearOptimizationService.LinearOptimizationEngine}
 * @param responses {string[][]}
 * @param headers {{[key: string]: number}}
 * @param params {{[name: string]: {name: string, min: number, max: number}}}
 * @param config {{idColumn: string}}
 */
function constrainCountPerShift(opt, responses, headers, params, config) {
  for (const [paramName, param] of Object.entries(params)) {
    const countConstraint = opt.addConstraint(param.min, param.max)
    const availColumnIndex = headers[paramName]
    const idColumnIndex = headers[config.idColumn]
    for (let row = 1; row < responses.length; ++row) {
      if (responses[row][availColumnIndex] !== 'Unavailable') {
        countConstraint.setCoefficient(getVariableName(responses[row][idColumnIndex], paramName), 1)
      }
    }
  }
}
/**
 * @param solution {LinearOptimizationService.LinearOptimizationSolution}
 * @param responses {string[][]}
 * @param headers {{[key: string]: number}}
 * @param params {{[name: string]: {name: string, min: number, max: number, outputRange: SpreadsheetApp.Range}}}
 * @param config {{idColumn: string, displayColumn: string, preferredGreen: boolean}}
 */
function writeSolution(solution, responses, headers, params, config) {
  // To enable us to write into multiple rows efficiently
  const shiftCounters = {}
  for (let row = 1; row < responses.length; ++row) {
    for (const [paramName, param] of Object.entries(params)) {
      const column = headers[paramName]
      const person = responses[row][headers[config.idColumn]]
      const personDisplayName = responses[row][headers[config.displayColumn]]
      // Logger.log({param: paramName, person: person, val: solution.getVariableValue(getVariableName(person, paramName))})
      if (responses[row][column] !== 'Unavailable' && solution.getVariableValue(getVariableName(person, paramName)) > 0) {
        let count = shiftCounters[paramName]
        if (count === undefined) {
          count = 0
          shiftCounters[paramName] = 0
        }
        const sheet = param.outputRange.getSheet()
        const nameRange = sheet.getRange(param.outputRange.getRow() + count, param.outputRange.getColumn())
        nameRange.setValue(personDisplayName)
        if (config.preferredGreen && responses[row][column] === 'Preferred') {
          nameRange.setFontColor('#47882b')
        } else {
          nameRange.setFontColor(null)
        }
        ++shiftCounters[paramName]
      }
    }
  }
}

function runSolver() {
  const params = collectParams()
  const responses = collectResponseArray()
  const opt = LinearOptimizationService.createEngine()
  const config = collectConfig()
  opt.setMaximization()
  const responseHeaders = mapHeaderToIndex(responses[0])

  makeAvailableVariables(opt, responses, responseHeaders, params, config)
  constrainHoursPerPerson(opt, responses, responseHeaders, params, config)
  constrainCountPerShift(opt, responses, responseHeaders, params, config)

  const solution = opt.solve(60)
  if (solution.isValid()) {
    writeSolution(solution, responses, responseHeaders, params, config)
    SpreadsheetApp.getUi().alert('Shifter says...', 'Shift scheduling done! ☺️', SpreadsheetApp.getUi().ButtonSet.OK)
  } else {
    SpreadsheetApp.getUi().alert('Shifter says...', 'Failed to solve linear program', SpreadsheetApp.getUi().ButtonSet.OK)
  }
}

