/* Swap chain digraph
- A swap chain (A-B-C-D) involves swapping person x on A with y on B, swapping person y on B with person z on C, and swapping person z on C with person v on D.
- Each link in the chain is made possible by a person on A being available for B and a person on B available for A.
*/

function formatRowToName(row: any[]): string {
    return `${row[2]} ${row[3]} ${row[5]}`
}

function querySwapsFrom(data: any[][], col: number): {col: number, fromPerson: string, toPerson: string, preference: number}[] {
    const availableRows: Map<number, number> = new Map();
    const assignedRows: Set<number> = new Set();
    const availableOut: Map<number, Set<string>> = new Map();
    const preferredOut: Map<number, Set<string>> = new Map();
    function fillOutMaps(row: number) {
        for (let otherCol = col + 1; otherCol <= 29; ++otherCol) {
            if (data[row][otherCol] === 'Available') {
                if (!availableOut.has(otherCol)) {
                    availableOut.set(otherCol, new Set());
                }
                availableOut.get(otherCol)!.add(formatRowToName(data[row]))
            }
            if (data[row][otherCol] === 'Preferred') {
                if (!preferredOut.has(otherCol)) {
                    preferredOut.set(otherCol, new Set());
                }
                preferredOut.get(otherCol)!.add(formatRowToName(data[row]))
            }
        }
    }
    for (let row = 1; row < data.length; ++row) {
        if (data[row][col] === 'Assigned') {
            assignedRows.add(row)
            fillOutMaps(row)
        } else if (data[row][col] === 'Preferred') {
            availableRows.set(row, 1)
        }
        else if (data[row][col] === 'Available') {
            availableRows.set(row, 0)
        }
    }
    // a swap with column X from keys(availableOut) union keys(preferredOut) is possible if there exists a row Y where data[Y][X] is assigned and Y is in availableRows or preferredRows
    function listPossibleSwapsToCol(col: number): {fromPerson: string, toPerson: string, preference: number}[] {
        const outwardList: {fromPerson: string, toPerson: string, preference: number}[] = []
        for (let row = 1; row < data.length; ++row) {
            if (data[row][col] === 'Assigned') {
                if (availableRows.has(row) && availableOut.has(col)) {
                    for (const fromPerson of availableOut.get(col)!) {
                        outwardList.push({fromPerson, toPerson: formatRowToName(data[row]), preference: availableRows.get(row)!})
                    }
                }
                if (availableRows.has(row) && preferredOut.has(col)) {
                    for (const fromPerson of preferredOut.get(col)!) {
                        outwardList.push({fromPerson, toPerson: formatRowToName(data[row]), preference: availableRows.get(row)! + 1})
                    }
                }
            }
        }
        return outwardList;
    }
    return Array.from(new Set([...availableOut.keys(), ...preferredOut.keys()]))
        .flatMap(col => listPossibleSwapsToCol(col).map(swap => ({...swap, col})))
}

function colToShiftName(col: number): string {
    const weekdays = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
    const shiftTypes = ['Cook', 'Clean']
    if (col >= 11 && col <= 24) {
        return `${weekdays[(col - 11) % 7]}${shiftTypes[Math.floor((col - 11) / 7)]}`
    }
    return ['SunLunchPrep', 'TueLunchPrep', 'ThuLunchPrep', 'MonChefHelper', 'ThuChefHelper'][col - 25]
}

// Shift-to-shift perspective
// A shift is connected to another shift when there is a pair of people that can swap between them
function generateShiftChainGraph() {
    const sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1YXLN4Z9BE0yVHfZ0xQxxuPTDJgwuPHq9d5lLl2RAjC0/edit')!.getSheetByName('Copy of Form Responses 1')!
    const sheetData = sheet.getDataRange().getValues()
    // column 11 to column 24 (both inclusive) are for the cook and clean shifts
    const connections: [string, string, [string, number][]][] = []
    const edgeMatrix: number[][] = []

    for (let col = 11; col <= 29; ++col) {
        const row = []
        for (let col2 = 11; col2 <= 29; ++col2) {
            row.push(0)
        }
        edgeMatrix.push(row)
    }

    for (let col = 11; col <= 29; ++col) {
        // This does not search leftward (to prevent redundant cycles)
        const swapsFromThis = querySwapsFrom(sheetData, col)
        const swapsByCol = new Map<number, Omit<typeof swapsFromThis[number], "col">[]>()
        for (const swap of swapsFromThis) {
            ++edgeMatrix[col - 11][swap.col - 11]
            ++edgeMatrix[swap.col - 11][col - 11]
            let colSwaps = swapsByCol.get(swap.col)
            if (colSwaps == undefined) {
                colSwaps = []
                swapsByCol.set(swap.col, colSwaps)
            }
            colSwaps.push(swap)
        }
        for (const [toCol, swaps] of swapsByCol.entries()) {
            connections.push([
                colToShiftName(col), colToShiftName(toCol),
                swaps.map(swap =>
                    [`${swap.fromPerson} <-> ${swap.toPerson} (${swap.preference})`, swap.preference])])
        }
    }
    function formatConnection(connection: typeof connections[number]): string {
        const penWidth = connection[2].map(conn => conn[1]).reduce((a, b) => a + b, 0)
        const label = connection[2].map(conn => conn[0]).join(', ')
        return `${connection[0]} -- ${connection[1]}[penwidth=${penWidth},weight=${penWidth},tooltip="${label}",label="${connection[2].length} swaps"];`
    }
    return {
        graph: `graph {\n${connections.map(formatConnection).join('\n')}\n}`,
        matrix: edgeMatrix
    }
}

function doGet() {
    return ContentService.createTextOutput(generateShiftChainGraph().graph).setMimeType(ContentService.MimeType.TEXT)
}

function fillAdjacencyMatrix() {
    const data = generateShiftChainGraph().matrix
    const sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1YXLN4Z9BE0yVHfZ0xQxxuPTDJgwuPHq9d5lLl2RAjC0/edit')!.getSheetByName('Swap Adjacency Matrix')!
    sheet.getRange(1, 25, data.length, data[0].length).setValues(data)
}

function* iteratePreferredSwitchers(from: number, to: number) {
    const sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/d/1YXLN4Z9BE0yVHfZ0xQxxuPTDJgwuPHq9d5lLl2RAjC0/edit')!.getSheetByName('Copy of Form Responses 1')!
    const sheetData = sheet.getDataRange().getValues()
    for (let row = 1; row < sheetData.length; ++row) {
        if (sheetData[row][from] === 'Assigned' && sheetData[row][to] === 'Preferred') {
            yield formatRowToName(sheetData[row])
        }
    }
}

function listPreferredSwitchers(from: number, to: number) {
    return Array.from(iteratePreferredSwitchers(from, to))
}