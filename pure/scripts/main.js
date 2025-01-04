// scripts/main.js

document.addEventListener("DOMContentLoaded", () => {
  const gridContainer = document.getElementById("grid-container")
  const rows = 40 // 데이터 행 수
  const cols = 30 // 데이터 열 수
  const cellData = {} // 셀 데이터를 저장할 객체

  // 열 헤더 생성 (A, B, C, ...)
  const headerRow = document.createElement("div")
  headerRow.classList.add("cell", "header-cell")
  gridContainer.appendChild(headerRow) // 빈 셀 (왼쪽 상단)

  for (let j = 0; j < cols; j++) {
    const headerCell = document.createElement("div")
    headerCell.classList.add("cell", "header-cell")
    headerCell.textContent = String.fromCharCode(65 + j) // A, B, C, ...
    gridContainer.appendChild(headerCell)
  }

  // 데이터 행 생성
  for (let i = 1; i <= rows; i++) {
    // 행 헤더 생성 (1, 2, 3, ...)
    const rowHeader = document.createElement("div")
    rowHeader.classList.add("cell", "header-cell")
    rowHeader.textContent = i
    gridContainer.appendChild(rowHeader)

    // 데이터 셀 생성
    for (let j = 0; j < cols; j++) {
      const cell = document.createElement("div")
      cell.classList.add("cell")
      cell.setAttribute("data-row", i)
      cell.setAttribute("data-col", j)
      cell.setAttribute("contenteditable", "false") // 초기에는 편집 불가
      gridContainer.appendChild(cell)
    }
  }

  gridContainer.addEventListener("click", (e) => {
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      enableEditing(target)
    }
  })

  // 편집 모드 활성화 함수
  function enableEditing(cell) {
    if (cell.querySelector("input")) return

    const row = cell.getAttribute("data-row")
    const col = cell.getAttribute("data-col")
    const cellKey = `${row}-${col}`
    const currentValue = cellData[cellKey] ?? ""

    cell.textContent = ""
    const input = document.createElement("input")
    input.type = "text"
    input.value = currentValue
    cell.appendChild(input)
    input.focus()

    input.addEventListener("blur", () => {
      const newValue = input.value.trim()
      cellData[cellKey] = newValue
      cell.textContent = cellData[cellKey]
    })

    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        input.blur()
      }
    })
  }
})
