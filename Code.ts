const objectiveTiers: {[tier: string]: {obj: number, green: boolean}} = {
  'Available': {obj: 0, green: false},
  'Preferred': {obj: 1, green: true},
  'Recommended': {obj: 3, green: true}
}

function onOpen() {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu('Shifter')
    .addItem('Validate', 'validateParams')
    .addItem('Solve', 'runSolver')
    .addToUi()
}

function mapHeaderToIndex(row: string[]): {[header: string]: number} {
  const paramHeaderToIndex: {[header: string]: number} = {}
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

    const expectedShifts: string[] = []
    for (let row = 1; row < params.length && params[row][0] != ''; ++row) {
      expectedShifts.push(params[row][paramHeaderToIndex['Shift']])
    }

    const deadlineCell = spreadsheet.getRangeByName('Deadline')
    if (deadlineCell != null) {
      if (!(deadlineCell.getValue() instanceof Date)) {
        deadlineCell.activate()
        throw new Error(`The selected deadline is not a valid date`)
      }

      const minOnTimePriorityCell = spreadsheet.getRangeByName('MinOnTimePriority')
      if (minOnTimePriorityCell == null)
        throw new Error(`You are missing the "MinOnTimePriority" named range, which specifies the priority given to workers who submit at the deadline. Conventionally, its value is somewhere around 0.5.`)
      if (typeof minOnTimePriorityCell.getValue() !== 'number') {
        minOnTimePriorityCell.activate()
        throw new Error(`The minimum priority for on-time submissions must be a number`)
      }

      const latePriorityCell = spreadsheet.getRangeByName('LatePriority')
      if (latePriorityCell == null)
        throw new Error(`You are missing the "LatePriority" named range, which specifies the priority given to workers who submit after the deadline. Conventionally, its value is somewhere around 0.`)
      if (typeof latePriorityCell.getValue() !== 'number') {
        latePriorityCell.activate()
        throw new Error(`The priority given to submissions after the deadline must be a number`)
      }
    }

    const additiveRange = spreadsheet.getRangeByName('Additive')
    if (additiveRange === null)
      throw new Error('You are missing the "Additive" named range, which specifies whether to add to an existing work schedule or create one from scratch')
    if (typeof additiveRange.getValue() !== 'boolean') {
      additiveRange.activate()
      throw new Error('The "Additive" named range must be a boolean (TRUE or FALSE)')
    }

    const responseData = responses.getDataRange().getValues()
    const missingShiftData = expectedShifts.filter(shift => !responseData[0].includes(shift))
    if (missingShiftData.length > 0) throw new Error(`You are missing response columns in ${constraintRange.getDisplayValue()} for the specified shifts: ${missingShiftData.join(", ")}`)
    const responseDataHeaders = mapHeaderToIndex(responseData[0])

    const knownCategories: Set<string> = new Set()

    // Check parameters
    for (let row = 1; params[row][0] != ''; ++row) {
      const minVal = params[row][paramHeaderToIndex['Minimum assigned']]
      const maxVal = params[row][paramHeaderToIndex['Maximum assigned']]
      const hours = params[row][paramHeaderToIndex['Hours']]
      const assignTo = params[row][paramHeaderToIndex['Write results to']]
      const category = params[row][paramHeaderToIndex['Category']]
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
      knownCategories.add(category)
    }
    
    if (!responseData[0].includes(idColumnRange.getDisplayValue()))
      throw new Error(`You are missing the "${idColumnRange.getDisplayValue()}" response column, which you have specified as the column with each person's unique ID`)
    if (!responseData[0].includes(displayColumnRange.getDisplayValue()))
      throw new Error(`You are missing the "${displayColumnRange.getDisplayValue()}" response column, which you have specified as the column with each person's display name in the schedule`)
    if (deadlineCell != null && !responseData[0].includes('Time'))
      throw new Error(`You are missing the "Time" response column, which is necessary because you have specified a submission deadline`)
    if (new Set(responseData[0]).size !== responseData[0].length)
      throw new Error(`The column headers of the responses sheet are not unique`)
    
    const resIdColumnIndex = responseDataHeaders[idColumnRange.getDisplayValue()]
    const resDisplayColumnIndex = responseDataHeaders[displayColumnRange.getDisplayValue()]
    const resReqCategoryColumnIndex = responseDataHeaders['Category']

    // TODO scale this beyond 0-1, like "Available", "Preferred", "Ideal"
    const validAvailability = new Set(['Unavailable', ...Object.keys(objectiveTiers)])
    const idSet = new Set()
    const nameSet = new Set()
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
      if (deadlineCell != null && !(responseData[row][responseDataHeaders['Time']] instanceof Date)) {
        responses.getRange(row + 1, responseDataHeaders['Time'] + 1).activate()
        throw new Error(`The selected worker submission time is missing or not a date`)
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

      const displayName = responseData[row][resDisplayColumnIndex].toString()
      if (nameSet.has(displayName)) {
        responses.getRange(row + 1, resDisplayColumnIndex + 1).activate()
        throw new Error(`The display name "${displayName}" is duplicated in the responses. Because of the mechanism in which the Additive feature reads the existing work schedule, all display names must be unique`)
      }
      nameSet.add(displayName)

      if (resReqCategoryColumnIndex != undefined) {
        const reqCategory: string|undefined = responseData[row][resReqCategoryColumnIndex]?.toString()
        if (reqCategory != null && reqCategory.length > 0) {
          const categories = reqCategory.split(/\s*,\s*/)
          const unknownCategory = categories.find(category => !knownCategories.has(category))
          if (unknownCategory !== undefined) {
            responses.getRange(row + 1, resReqCategoryColumnIndex + 1).activate()
            throw new Error(`The mentioned category "${unknownCategory}" in the selected cell is unknown. Is it a typo?`)
          }
        }
      }
    }
    
    if (additiveRange.getValue() === true) {
      for (let row = 1; params[row][0] != ''; ++row) {
        let assignTo = config.getRange(params[row][paramHeaderToIndex['Write results to']])
        while (nameSet.has(assignTo?.getValue())) {
          assignTo = config.getRange(assignTo.getRow() + 1, assignTo.getColumn())
        }
        if (/[a-zA-Z]/.test(assignTo.getValue()?.toString() ?? '')) {
          assignTo.activate()
          throw new Error(`${JSON.stringify(assignTo.getValue())} exists in the current schedule but not in the availability responses`)
        }
      }
    }

    SpreadsheetApp.getUi().alert('Shifter says...', `Looks good! 👍`, SpreadsheetApp.getUi().ButtonSet.OK)
  } catch (e) {
    SpreadsheetApp.getUi().alert('Shifter has encountered a problem', `${e?.stack}`, SpreadsheetApp.getUi().ButtonSet.OK)
  }
}

interface ShiftParams {
  name: string
  category: string
  min: number // int
  max: number // int
  hours: number // float
  outputRange: GoogleAppsScript.Spreadsheet.Range
}

// Precondition: Already validated
function collectParams() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const outputSheet = spreadsheet.getSheetByName('Shifter')!
  const paramsRange = spreadsheet.getRangeByName('Shifts')!
  const paramsData = paramsRange.getDisplayValues()
  const paramHeaderToIndex = mapHeaderToIndex(paramsData[0])
  const params: {[shiftName: string]: ShiftParams} = {}
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

function collectResponseArray(): any[][] {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    const constraintRange = spreadsheet.getRangeByName('AvailabilityResponses')!
    const responses = spreadsheet.getSheetByName(constraintRange.getDisplayValue())!
    return responses.getDataRange().getValues()
}

function getVariableName(person: string, shift: string) {
  return `${person}_${shift}`
}

interface Config {
  idColumn: string
  displayColumn: string
  preferredGreen: boolean
  minHours: number
  maxHours: number
  additive: boolean
  timeScale?: {
    deadline: Date,
    minOnTimePriority: number,
    latePriority: number
  }
}

function collectConfig() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const idColumn = spreadsheet.getRangeByName('IdColumn')!.getDisplayValue()
  const displayColumn = spreadsheet.getRangeByName('DisplayColumn')!.getDisplayValue()
  const preferredGreenStr = spreadsheet.getRangeByName('PreferredGreen')!.getDisplayValue()
  const preferredGreen = !['0', '', 'FALSE'].includes(preferredGreenStr.toString().toUpperCase())
  const minHours = spreadsheet.getRangeByName('MinHours')!.getValue()
  const maxHours = spreadsheet.getRangeByName('MaxHours')!.getValue()
  const additive = spreadsheet.getRangeByName('Additive')!.getValue()
  const config: Config = {idColumn, displayColumn, preferredGreen, minHours, maxHours, additive}

  const deadlineCell = spreadsheet.getRangeByName('Deadline')
  if (deadlineCell != null) {
    config.timeScale = {
      deadline: deadlineCell.getValue() as Date,
      minOnTimePriority: spreadsheet.getRangeByName('MinOnTimePriority')!.getValue() as number,
      latePriority: spreadsheet.getRangeByName('LatePriority')!.getValue() as number
    }
  }
  return config
}

function readCurrentAssignments(problem: ShifterProblem): {[shiftName: string]: string[]} {
  const displayNameMapping = Object.fromEntries(
    problem.responses.map(row => 
      [row[problem.responseHeaders[problem.config.displayColumn]], row[problem.responseHeaders[problem.config.idColumn]]]))
  const currentAssignments: {[shiftName: string]: string[]} = {}

  for (const [shiftName, params] of Object.entries(problem.shifts)) {
    currentAssignments[shiftName] = []
    for (let item = params.outputRange; item?.getValue() in displayNameMapping; item = item.getSheet().getRange(item.getRow() + 1, item.getColumn())) {
      currentAssignments[shiftName].push(displayNameMapping[item.getValue()])
    }
  }
  return currentAssignments
}

class ShifterProblem {
  shifts: {[shiftName: string]: ShiftParams}
  responses: any[][]
  currentAssignments?: {[task: string]: string[]}
  opt: GoogleAppsScript.Optimization.LinearOptimizationEngine
  config: Config
  responseHeaders: {[header: string]: number}
  earliestTime?: number = undefined

  constructor() {
    this.shifts = collectParams()
    this.responses = collectResponseArray()
    this.opt = LinearOptimizationService.createEngine()
    this.config = collectConfig()
    this.opt.setMaximization()
    this.responseHeaders = mapHeaderToIndex(this.responses[0])
    if (this.config.additive) {
      this.currentAssignments = readCurrentAssignments(this)
    }
  }

  makeAvailabilityVariables() {
    // For all workers
    for (let row = 1; row < this.responses.length; ++row) {
      // For all shifts
      for (const paramName of Object.keys(this.shifts)) {
        const response = this.responses[row][this.responseHeaders[paramName]]
        const shiftCategory = this.shifts[paramName].category
        const categoryPermitted = this.isCategoryPermitted(this.responses[row], shiftCategory)
        if (categoryPermitted) {
          const varName = getVariableName(this.responses[row][this.responseHeaders[this.config.idColumn]], paramName)
          if (response !== 'Unavailable') {
            this.opt.addVariable(varName, 0, 1, LinearOptimizationService.VariableType.INTEGER)
          }
          const tier = objectiveTiers[response]
          if (tier !== undefined) {
            if (this.config.timeScale == null) {
              this.opt.setObjectiveCoefficient(varName, tier.obj)
            } else {
              const submissionTime = this.responses[row][this.responseHeaders['Time']]
              this.opt.setObjectiveCoefficient(varName, tier.obj * this.scaleSubmissionTimeToWeight(submissionTime))
            }
          }
        }
      }
    }
  }

  private isCategoryPermitted(responseRow: any[], category: string): boolean {
    const categoryColumnIndex = this.responseHeaders['Category']
    if (categoryColumnIndex == undefined) return true
    const reqCategory = responseRow[categoryColumnIndex]
    if (reqCategory == undefined || reqCategory === '') return true
    return reqCategory.includes(category)
  }

  private scaleSubmissionTimeToWeight(time: Date) {
    if (this.config.timeScale == null) return 1
    // cached
    this.earliestTime ??= this.responses.slice(1)
      .map(row => row[this.responseHeaders['Time']].getTime())
      .reduce((a, b) => Math.min(a, b))
    const deadline = this.config.timeScale.deadline.getTime()
    if (time.getTime() > deadline) return this.config.timeScale.latePriority

    // (earliest) 0 -> 1 (latest)
    // (earliest) 1 -> minOnTimePriority (latest)
    return 1 - (time.getTime() - this.earliestTime!) / (deadline - this.earliestTime!) * (1 - this.config.timeScale.minOnTimePriority)
  }

  constrainHoursPerPerson() {
    for (let row = 1; row < this.responses.length; ++row) {
      const shiftConstraint = this.opt.addConstraint(this.config.minHours, this.config.maxHours)
      for (const shiftName of Object.keys(this.shifts)) {
        const response = this.responses[row][this.responseHeaders[shiftName]]
        const varName = getVariableName(this.responses[row][this.responseHeaders[this.config.idColumn]], shiftName)
        if (response !== 'Unavailable' && this.isCategoryPermitted(this.responses[row], this.shifts[shiftName].category)) {
          shiftConstraint.setCoefficient(varName, this.shifts[shiftName].hours)
        }
      }
    }
  }

  constrainCountPerShift() {
    const idColumnIndex = this.responseHeaders[this.config.idColumn]
    for (const [shiftName, param] of Object.entries(this.shifts)) {
      const countConstraint = this.opt.addConstraint(param.min, param.max)
      const availColumnIndex = this.responseHeaders[shiftName]
      for (let row = 1; row < this.responses.length; ++row) {
        if (this.responses[row][availColumnIndex] !== 'Unavailable' && this.isCategoryPermitted(this.responses[row], param.category)) {
          countConstraint.setCoefficient(getVariableName(this.responses[row][idColumnIndex], shiftName), 1)
        }
      }
    }
  }

  constrainExistingAssignmentsIfAdditive() {
    if (this.currentAssignments) {
      // We opt for a single constraint where the sum of all the existing assignments' variables should be their count
      const existingAssignmentCount = Object.values(this.currentAssignments).reduce((a, b) => a + b.length, 0)
      const existingConstraint = this.opt.addConstraint(existingAssignmentCount, existingAssignmentCount)
      for (const [shiftName, ids] of Object.entries(this.currentAssignments)) {
        for (const id of ids) {
          existingConstraint.setCoefficient(getVariableName(id, shiftName), 1)
        }
      }
    }
  }

  writeSolution(solution: GoogleAppsScript.Optimization.LinearOptimizationSolution) {
    // To enable us to write into multiple rows efficiently
    const shiftCounters = {}
    for (let row = 1; row < this.responses.length; ++row) {
      for (const [paramName, param] of Object.entries(this.shifts)) {
        const column = this.responseHeaders[paramName]
        const person = this.responses[row][this.responseHeaders[this.config.idColumn]]
        const personDisplayName = this.responses[row][this.responseHeaders[this.config.displayColumn]]
        // Logger.log({param: paramName, person: person, val: solution.getVariableValue(getVariableName(person, paramName))})
        if (this.responses[row][column] !== 'Unavailable' && this.isCategoryPermitted(this.responses[row], param.category) && solution.getVariableValue(getVariableName(person, paramName)) > 0) {
          let count = shiftCounters[paramName]
          if (count === undefined) {
            count = 0
            shiftCounters[paramName] = 0
          }
          const sheet = param.outputRange.getSheet()
          const nameRange = sheet.getRange(param.outputRange.getRow() + count, param.outputRange.getColumn())
          nameRange.setValue(personDisplayName)
          if (this.config.preferredGreen && objectiveTiers[this.responses[row][column]]?.green) {
            nameRange.setFontColor('#47882b')
          } else {
            nameRange.setFontColor(null)
          }
          ++shiftCounters[paramName]
        }
      }
    }
  }
}


function runSolver() {
  const prob = new ShifterProblem()
  prob.makeAvailabilityVariables()
  prob.constrainHoursPerPerson()
  prob.constrainCountPerShift()
  prob.constrainExistingAssignmentsIfAdditive()

  const solution = prob.opt.solve(60)
  if (solution.isValid()) {
    prob.writeSolution(solution)
    SpreadsheetApp.getUi().alert('Shifter says...', 'Shift scheduling done! ☺️', SpreadsheetApp.getUi().ButtonSet.OK)
  } else {
    SpreadsheetApp.getUi().alert('Shifter says...', 'Failed to solve linear program', SpreadsheetApp.getUi().ButtonSet.OK)
  }
}

// TEST CASES
// Validate: Check duplicate display names
// Validate: Check presence of "Additive" named range and that it is a boolean
// Validate: Check validity of all display names in the output section if Additive
// Validate: Unknown category name in Category column
