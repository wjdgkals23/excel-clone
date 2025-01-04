// scripts/main.js

document.addEventListener("DOMContentLoaded", () => {
  const gridContainer = document.getElementById("grid-container")
  const rows = 20 // 데이터 행 수
  const cols = 10 // 데이터 열 수
  const cellData = {} // 셀 데이터를 저장할 객체
  let selectedCells = new Set() // 선택된 셀을 추적
  let currentCell = null // 현재 선택된 셀
  let isResizing = false
  let currentResizer = null
  let startX, startWidth
  let isRowResizing = false
  let currentRowResizer = null
  let startY, startHeight
  let dragStartCell = null
  const contextMenu = document.getElementById("context-menu")
  let contextTargetCell = null

  // 그리드 생성 함수
  function renderGrid() {
    showLoading()
    setTimeout(() => {
      // 비동기적으로 처리하는 것처럼 시뮬레이션
      gridContainer.innerHTML = ""
      const fragment = document.createDocumentFragment()

      // 열 헤더 생성 (A, B, C, ...)
      const emptyHeader = document.createElement("div")
      emptyHeader.classList.add("cell", "header-cell")
      fragment.appendChild(emptyHeader) // 왼쪽 상단 빈 셀

      for (let j = 0; j < cols; j++) {
        const headerCell = document.createElement("div")
        headerCell.classList.add("cell", "header-cell")
        headerCell.textContent = String.fromCharCode(65 + j) // A, B, C, ...
        fragment.appendChild(headerCell)
      }

      // 데이터 행 생성
      for (let i = 1; i <= rows; i++) {
        // 행 헤더 생성 (1, 2, 3, ...)
        const rowHeader = document.createElement("div")
        rowHeader.classList.add("cell", "header-cell", "row-header")
        rowHeader.textContent = i
        fragment.appendChild(rowHeader)

        // 데이터 셀 생성
        for (let j = 0; j < cols; j++) {
          const cell = document.createElement("div")
          cell.classList.add("cell")
          cell.setAttribute("data-row", i)
          cell.setAttribute("data-col", j)
          cell.setAttribute("contenteditable", "false") // 초기에는 편집 불가
          cell.setAttribute("draggable", "true") // 드래그 가능하도록 설정
          cell.textContent =
            cellData[`${i}-${j}`] || `${String.fromCharCode(65 + j)}${i}` // 초기 값 표시
          fragment.appendChild(cell)
        }
      }

      gridContainer.appendChild(fragment)
      hideLoading()
      addColumnResizers() // 컬럼 리사이저 추가
      addRowResizers() // 로우 리사이저 추가
      enableDragAndDrop() // 드래그 앤 드롭 활성화
    }, 500) // 0.5초 후에 렌더링 완료
  }

  // 로딩 스피너 제어 함수
  function showLoading() {
    document.getElementById("loading-spinner").style.display = "block"
  }

  function hideLoading() {
    document.getElementById("loading-spinner").style.display = "none"
  }

  // 셀 클릭 시 선택 토글 및 현재 셀 설정
  gridContainer.addEventListener("click", (e) => {
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      currentCell = target
      toggleSelection(target)
    }
  })

  // 셀 더블 클릭 시 편집 모드 활성화
  gridContainer.addEventListener("dblclick", (e) => {
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      enableEditing(target)
    }
  })

  // 셀 편집 활성화 함수
  function enableEditing(cell) {
    if (cell.querySelector("input")) return // 이미 편집 중인 경우

    const currentValue = cell.textContent
    cell.textContent = ""
    const input = document.createElement("input")
    input.type = "text"
    input.value = currentValue
    cell.appendChild(input)
    input.focus()

    // 입력 완료 시 저장
    input.addEventListener("blur", () => {
      const newValue = input.value.trim()
      cell.textContent =
        newValue ||
        `${String.fromCharCode(
          65 + parseInt(cell.getAttribute("data-col"))
        )}${cell.getAttribute("data-row")}`
    })

    // Enter 키로 저장
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        input.blur()
      }
    })
  }

  // 셀 선택 토글 함수
  function toggleSelection(cell) {
    if (selectedCells.has(cell)) {
      cell.classList.remove("selected")
      selectedCells.delete(cell)
    } else {
      cell.classList.add("selected")
      selectedCells.add(cell)
    }
  }

  // 선택 해제 함수
  function clearSelection() {
    selectedCells.forEach((cell) => cell.classList.remove("selected"))
    selectedCells.clear()
  }

  // 도구 모음 버튼 기능 구현
  const bgColorBtn = document.getElementById("bg-color-btn")
  bgColorBtn.addEventListener("click", () => {
    changeBackgroundColor(Array.from(selectedCells))
  })

  const textColorBtn = document.getElementById("text-color-btn")
  textColorBtn.addEventListener("click", () => {
    changeTextColor(Array.from(selectedCells))
  })

  const textAlignLeft = document.getElementById("text-align-left")
  textAlignLeft.addEventListener("click", () => {
    applyTextAlign("left", Array.from(selectedCells))
  })

  const textAlignCenter = document.getElementById("text-align-center")
  textAlignCenter.addEventListener("click", () => {
    applyTextAlign("center", Array.from(selectedCells))
  })

  const textAlignRight = document.getElementById("text-align-right")
  textAlignRight.addEventListener("click", () => {
    applyTextAlign("right", Array.from(selectedCells))
  })

  const mergeCellsBtn = document.getElementById("merge-cells-btn")
  mergeCellsBtn.addEventListener("click", () => {
    mergeSelectedCells()
  })

  const themeToggleBtn = document.getElementById("theme-toggle-btn")
  themeToggleBtn.addEventListener("click", () => {
    document.body.classList.toggle("dark-mode")
  })

  const fontSelectBtn = document.getElementById("font-select-btn")
  fontSelectBtn.addEventListener("click", () => {
    const font = prompt(
      "원하는 폰트를 입력하세요 (예: Arial, Times New Roman):",
      "Arial"
    )
    if (font) {
      selectedCells.forEach((cell) => {
        cell.style.fontFamily = font
      })
    }
  })

  const boldBtn = document.getElementById("bold-btn")
  boldBtn.addEventListener("click", () => {
    selectedCells.forEach((cell) => {
      cell.style.fontWeight =
        cell.style.fontWeight === "bold" ? "normal" : "bold"
    })
  })

  const italicBtn = document.getElementById("italic-btn")
  italicBtn.addEventListener("click", () => {
    selectedCells.forEach((cell) => {
      cell.style.fontStyle =
        cell.style.fontStyle === "italic" ? "normal" : "italic"
    })
  })

  const alignLeftBtn = document.getElementById("align-left-btn")
  alignLeftBtn.addEventListener("click", () => {
    applyTextAlignToCells("left", Array.from(selectedCells))
  })

  const alignCenterBtn = document.getElementById("align-center-btn")
  alignCenterBtn.addEventListener("click", () => {
    applyTextAlignToCells("center", Array.from(selectedCells))
  })

  const alignRightBtn = document.getElementById("align-right-btn")
  alignRightBtn.addEventListener("click", () => {
    applyTextAlignToCells("right", Array.from(selectedCells))
  })

  // 텍스트 정렬 적용 함수
  function applyTextAlign(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  function applyTextAlignToCells(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  // 배경색 변경 함수
  function changeBackgroundColor(cells) {
    const color = prompt(
      "원하는 배경색을 입력하세요 (예: #ff0000 또는 red):",
      "#ffff00"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.backgroundColor = color
      })
    }
  }

  // 글자색 변경 함수
  function changeTextColor(cells) {
    const color = prompt(
      "원하는 글자색을 입력하세요 (예: #ff0000 또는 red):",
      "#000000"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.color = color
      })
    }
  }

  // 셀 병합 함수 (기본적인 병합)
  function mergeSelectedCells(cells = Array.from(selectedCells)) {
    if (cells.length <= 1) {
      alert("병합할 셀을 두 개 이상 선택하세요.")
      return
    }

    // 선택된 셀들을 배열로 변환
    const cellsArray = cells
    // 첫 번째 셀을 기준으로 병합
    const firstCell = cellsArray[0]
    const firstRow = parseInt(firstCell.getAttribute("data-row"))
    const firstCol = parseInt(firstCell.getAttribute("data-col"))

    // 병합할 셀의 범위 계산
    let minRow = firstRow,
      maxRow = firstRow
    let minCol = firstCol,
      maxCol = firstCol

    cellsArray.forEach((cell) => {
      const row = parseInt(cell.getAttribute("data-row"))
      const col = parseInt(cell.getAttribute("data-col"))
      if (row < minRow) minRow = row
      if (row > maxRow) maxRow = row
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    })

    // 병합 범위에 있는 셀들을 숨기고 첫 번째 셀에 rowspan과 colspan 적용
    for (let i = minRow; i <= maxRow; i++) {
      for (let j = minCol; j <= maxCol; j++) {
        const selector = `.cell[data-row='${i}'][data-col='${j}']`
        const cell = gridContainer.querySelector(selector)
        if (cell && cell !== firstCell) {
          cell.style.display = "none"
        }
      }
    }

    // 첫 번째 셀에 rowspan과 colspan 적용
    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1
    firstCell.style.gridRowEnd = `span ${rowspan}`
    firstCell.style.gridColumnEnd = `span ${colspan}`

    // 선택 해제
    clearSelection()
  }

  // 컬럼 리사이징 핸들러 추가
  function addColumnResizers() {
    const headers = gridContainer.querySelectorAll(".header-cell")
    headers.forEach((header) => {
      if (header.classList.contains("row-header")) return // 행 헤더는 제외
      const resizer = document.createElement("div")
      resizer.classList.add("column-resizer")
      header.appendChild(resizer)

      resizer.addEventListener("mousedown", (e) => {
        isResizing = true
        currentResizer = header
        startX = e.clientX
        startWidth = header.offsetWidth
        document.body.style.cursor = "col-resize"
        e.preventDefault()
      })
    })
  }

  // 로우 리사이징 핸들러 추가
  function addRowResizers() {
    const rowHeaders = gridContainer.querySelectorAll(".row-header")
    rowHeaders.forEach((header) => {
      const resizer = document.createElement("div")
      resizer.classList.add("row-resizer")
      header.appendChild(resizer)

      resizer.addEventListener("mousedown", (e) => {
        isRowResizing = true
        currentRowResizer = header
        startY = e.clientY
        startHeight = header.offsetHeight
        document.body.style.cursor = "row-resize"
        e.preventDefault()
      })
    })
  }

  // 컬럼 리사이징 이벤트 처리
  document.addEventListener("mousemove", (e) => {
    if (isResizing && currentResizer) {
      const dx = e.clientX - startX
      const newWidth = startWidth + dx
      if (newWidth > 50) {
        // 최소 너비 설정
        currentResizer.style.width = `${newWidth}px`
        const colIndex =
          Array.from(currentResizer.parentNode.parentNode.children).indexOf(
            currentResizer.parentNode
          ) - 1 // 첫 셀은 빈 헤더
        const cells = gridContainer.querySelectorAll(
          `.cell[data-col='${colIndex}']`
        )
        cells.forEach((cell) => {
          cell.style.width = `${newWidth}px`
        })
      }
    }

    if (isRowResizing && currentRowResizer) {
      const dy = e.clientY - startY
      const newHeight = startHeight + dy
      if (newHeight > 20) {
        // 최소 높이 설정
        currentRowResizer.style.height = `${newHeight}px`
        const rowIndex = parseInt(currentRowResizer.textContent) - 1
        const cells = gridContainer.querySelectorAll(
          `.cell[data-row='${rowIndex + 1}']`
        )
        cells.forEach((cell) => {
          cell.style.height = `${newHeight}px`
        })
      }
    }
  })

  // 리사이징 종료 이벤트 처리
  document.addEventListener("mouseup", () => {
    if (isResizing) {
      isResizing = false
      currentResizer = null
      document.body.style.cursor = "default"
    }
    if (isRowResizing) {
      isRowResizing = false
      currentRowResizer = null
      document.body.style.cursor = "default"
    }
  })

  // 드래그 앤 드롭 활성화 함수
  function enableDragAndDrop() {
    const cells = gridContainer.querySelectorAll(".cell:not(.header-cell)")
    cells.forEach((cell) => {
      cell.addEventListener("dragstart", (e) => {
        dragStartCell = cell
        e.dataTransfer.setData(
          "text/plain",
          `${cell.getAttribute("data-row")}-${cell.getAttribute("data-col")}`
        )
        cell.classList.add("dragging")
      })

      cell.addEventListener("dragend", () => {
        dragStartCell.classList.remove("dragging")
      })

      cell.addEventListener("dragover", (e) => {
        e.preventDefault()
        cell.classList.add("drag-over")
      })

      cell.addEventListener("dragleave", () => {
        cell.classList.remove("drag-over")
      })

      cell.addEventListener("drop", (e) => {
        e.preventDefault()
        cell.classList.remove("drag-over")
        const data = e.dataTransfer.getData("text/plain")
        const [startRow, startCol] = data.split("-").map(Number)
        const [endRow, endCol] = [
          cell.getAttribute("data-row"),
          cell.getAttribute("data-col"),
        ].map(Number)

        // 셀 데이터 교환
        const startCellKey = `${startRow}-${startCol}`
        const endCellKey = `${endRow}-${endCol}`
        const temp =
          cellData[startCellKey] ||
          `${String.fromCharCode(65 + startCol)}${startRow}`
        cellData[startCellKey] =
          cellData[endCellKey] || `${String.fromCharCode(65 + endCol)}${endRow}`
        cellData[endCellKey] = temp

        // 그리드 재렌더링
        renderGrid()
      })
    })
  }

  // 컨텍스트 메뉴 표시
  gridContainer.addEventListener("contextmenu", (e) => {
    e.preventDefault()
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      contextTargetCell = target
      showContextMenu(e.pageX, e.pageY)
    }
  })

  // 컨텍스트 메뉴 숨기기
  document.addEventListener("click", () => {
    hideContextMenu()
  })

  // 컨텍스트 메뉴 항목 클릭 처리
  contextMenu.addEventListener("click", (e) => {
    const action = e.target.id
    switch (action) {
      case "context-edit":
        enableEditing(contextTargetCell)
        break
      case "context-bg-color":
        changeBackgroundColor([contextTargetCell])
        break
      case "context-text-color":
        changeTextColor([contextTargetCell])
        break
      case "context-text-align-left":
        applyTextAlignToCells("left", [contextTargetCell])
        break
      case "context-text-align-center":
        applyTextAlignToCells("center", [contextTargetCell])
        break
      case "context-text-align-right":
        applyTextAlignToCells("right", [contextTargetCell])
        break
      case "context-merge-cells":
        mergeSelectedCells([contextTargetCell])
        break
      default:
        break
    }
    hideContextMenu()
  })

  // 컨텍스트 메뉴 표시 함수
  function showContextMenu(x, y) {
    contextMenu.style.top = `${y}px`
    contextMenu.style.left = `${x}px`
    contextMenu.style.display = "block"
  }

  // 컨텍스트 메뉴 숨기기 함수
  function hideContextMenu() {
    contextMenu.style.display = "none"
  }

  // 단축키 이벤트 처리
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.metaKey) {
      // Ctrl 또는 Command 키 조합
      switch (e.key.toLowerCase()) {
        case "b": // Ctrl+B: 배경색 변경
          e.preventDefault()
          changeBackgroundColor(Array.from(selectedCells))
          break
        case "c": // Ctrl+C: 글자색 변경
          e.preventDefault()
          changeTextColor(Array.from(selectedCells))
          break
        case "l": // Ctrl+L: 왼쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("left", Array.from(selectedCells))
          break
        case "e": // Ctrl+E: 가운데 정렬
          e.preventDefault()
          applyTextAlignToCells("center", Array.from(selectedCells))
          break
        case "r": // Ctrl+R: 오른쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("right", Array.from(selectedCells))
          break
        default:
          break
      }
    }

    // Esc 키로 선택 해제
    if (e.key === "Escape") {
      clearSelection()
    }
  })

  // 배경색 변경 함수
  function changeBackgroundColor(cells) {
    const color = prompt(
      "원하는 배경색을 입력하세요 (예: #ff0000 또는 red):",
      "#ffff00"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.backgroundColor = color
      })
    }
  }

  // 글자색 변경 함수
  function changeTextColor(cells) {
    const color = prompt(
      "원하는 글자색을 입력하세요 (예: #ff0000 또는 red):",
      "#000000"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.color = color
      })
    }
  }

  // 텍스트 정렬 함수
  function applyTextAlignToCells(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  // 셀 병합 함수
  function mergeSelectedCells(cells = Array.from(selectedCells)) {
    if (cells.length <= 1) {
      alert("병합할 셀을 두 개 이상 선택하세요.")
      return
    }

    // 선택된 셀들을 배열로 변환
    const cellsArray = cells
    // 첫 번째 셀을 기준으로 병합
    const firstCell = cellsArray[0]
    const firstRow = parseInt(firstCell.getAttribute("data-row"))
    const firstCol = parseInt(firstCell.getAttribute("data-col"))

    // 병합할 셀의 범위 계산
    let minRow = firstRow,
      maxRow = firstRow
    let minCol = firstCol,
      maxCol = firstCol

    cellsArray.forEach((cell) => {
      const row = parseInt(cell.getAttribute("data-row"))
      const col = parseInt(cell.getAttribute("data-col"))
      if (row < minRow) minRow = row
      if (row > maxRow) maxRow = row
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    })

    // 병합 범위에 있는 셀들을 숨기고 첫 번째 셀에 rowspan과 colspan 적용
    for (let i = minRow; i <= maxRow; i++) {
      for (let j = minCol; j <= maxCol; j++) {
        const selector = `.cell[data-row='${i}'][data-col='${j}']`
        const cell = gridContainer.querySelector(selector)
        if (cell && cell !== firstCell) {
          cell.style.display = "none"
        }
      }
    }

    // 첫 번째 셀에 rowspan과 colspan 적용
    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1
    firstCell.style.gridRowEnd = `span ${rowspan}`
    firstCell.style.gridColumnEnd = `span ${colspan}`

    // 선택 해제
    clearSelection()
  }

  // 컬럼 리사이징 이벤트 처리
  document.addEventListener("mousemove", (e) => {
    if (isResizing && currentResizer) {
      const dx = e.clientX - startX
      const newWidth = startWidth + dx
      if (newWidth > 50) {
        // 최소 너비 설정
        currentResizer.style.width = `${newWidth}px`
        const colIndex =
          Array.from(currentResizer.parentNode.parentNode.children).indexOf(
            currentResizer.parentNode
          ) - 1 // 첫 셀은 빈 헤더
        const cells = gridContainer.querySelectorAll(
          `.cell[data-col='${colIndex}']`
        )
        cells.forEach((cell) => {
          cell.style.width = `${newWidth}px`
        })
      }
    }

    if (isRowResizing && currentRowResizer) {
      const dy = e.clientY - startY
      const newHeight = startHeight + dy
      if (newHeight > 20) {
        // 최소 높이 설정
        currentRowResizer.style.height = `${newHeight}px`
        const rowIndex = parseInt(currentRowResizer.textContent) - 1
        const cells = gridContainer.querySelectorAll(
          `.cell[data-row='${rowIndex + 1}']`
        )
        cells.forEach((cell) => {
          cell.style.height = `${newHeight}px`
        })
      }
    }
  })

  // 리사이징 종료 이벤트 처리
  document.addEventListener("mouseup", () => {
    if (isResizing) {
      isResizing = false
      currentResizer = null
      document.body.style.cursor = "default"
    }
    if (isRowResizing) {
      isRowResizing = false
      currentRowResizer = null
      document.body.style.cursor = "default"
    }
  })

  // 드래그 앤 드롭 활성화 함수
  function enableDragAndDrop() {
    const cells = gridContainer.querySelectorAll(".cell:not(.header-cell)")
    cells.forEach((cell) => {
      cell.addEventListener("dragstart", (e) => {
        dragStartCell = cell
        e.dataTransfer.setData(
          "text/plain",
          `${cell.getAttribute("data-row")}-${cell.getAttribute("data-col")}`
        )
        cell.classList.add("dragging")
      })

      cell.addEventListener("dragend", () => {
        dragStartCell.classList.remove("dragging")
      })

      cell.addEventListener("dragover", (e) => {
        e.preventDefault()
        cell.classList.add("drag-over")
      })

      cell.addEventListener("dragleave", () => {
        cell.classList.remove("drag-over")
      })

      cell.addEventListener("drop", (e) => {
        e.preventDefault()
        cell.classList.remove("drag-over")
        const data = e.dataTransfer.getData("text/plain")
        const [startRow, startCol] = data.split("-").map(Number)
        const [endRow, endCol] = [
          cell.getAttribute("data-row"),
          cell.getAttribute("data-col"),
        ].map(Number)

        // 셀 데이터 교환
        const startCellKey = `${startRow}-${startCol}`
        const endCellKey = `${endRow}-${endCol}`
        const temp =
          cellData[startCellKey] ||
          `${String.fromCharCode(65 + startCol)}${startRow}`
        cellData[startCellKey] =
          cellData[endCellKey] || `${String.fromCharCode(65 + endCol)}${endRow}`
        cellData[endCellKey] = temp

        // 그리드 재렌더링
        renderGrid()
      })
    })
  }

  // 컨텍스트 메뉴 표시
  gridContainer.addEventListener("contextmenu", (e) => {
    e.preventDefault()
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      contextTargetCell = target
      showContextMenu(e.pageX, e.pageY)
    }
  })

  // 컨텍스트 메뉴 숨기기
  document.addEventListener("click", () => {
    hideContextMenu()
  })

  // 컨텍스트 메뉴 항목 클릭 처리
  contextMenu.addEventListener("click", (e) => {
    const action = e.target.id
    switch (action) {
      case "context-edit":
        enableEditing(contextTargetCell)
        break
      case "context-bg-color":
        changeBackgroundColor([contextTargetCell])
        break
      case "context-text-color":
        changeTextColor([contextTargetCell])
        break
      case "context-text-align-left":
        applyTextAlignToCells("left", [contextTargetCell])
        break
      case "context-text-align-center":
        applyTextAlignToCells("center", [contextTargetCell])
        break
      case "context-text-align-right":
        applyTextAlignToCells("right", [contextTargetCell])
        break
      case "context-merge-cells":
        mergeSelectedCells([contextTargetCell])
        break
      default:
        break
    }
    hideContextMenu()
  })

  // 컨텍스트 메뉴 표시 함수
  function showContextMenu(x, y) {
    contextMenu.style.top = `${y}px`
    contextMenu.style.left = `${x}px`
    contextMenu.style.display = "block"
  }

  // 컨텍스트 메뉴 숨기기 함수
  function hideContextMenu() {
    contextMenu.style.display = "none"
  }

  // 단축키 이벤트 처리
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.metaKey) {
      // Ctrl 또는 Command 키 조합
      switch (e.key.toLowerCase()) {
        case "b": // Ctrl+B: 배경색 변경
          e.preventDefault()
          changeBackgroundColor(Array.from(selectedCells))
          break
        case "c": // Ctrl+C: 글자색 변경
          e.preventDefault()
          changeTextColor(Array.from(selectedCells))
          break
        case "l": // Ctrl+L: 왼쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("left", Array.from(selectedCells))
          break
        case "e": // Ctrl+E: 가운데 정렬
          e.preventDefault()
          applyTextAlignToCells("center", Array.from(selectedCells))
          break
        case "r": // Ctrl+R: 오른쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("right", Array.from(selectedCells))
          break
        default:
          break
      }
    }

    // Esc 키로 선택 해제
    if (e.key === "Escape") {
      clearSelection()
    }
  })

  // 배경색 변경 함수
  function changeBackgroundColor(cells) {
    const color = prompt(
      "원하는 배경색을 입력하세요 (예: #ff0000 또는 red):",
      "#ffff00"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.backgroundColor = color
      })
    }
  }

  // 글자색 변경 함수
  function changeTextColor(cells) {
    const color = prompt(
      "원하는 글자색을 입력하세요 (예: #ff0000 또는 red):",
      "#000000"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.color = color
      })
    }
  }

  // 텍스트 정렬 함수
  function applyTextAlignToCells(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  // 셀 병합 함수
  function mergeSelectedCells(cells = Array.from(selectedCells)) {
    if (cells.length <= 1) {
      alert("병합할 셀을 두 개 이상 선택하세요.")
      return
    }

    // 선택된 셀들을 배열로 변환
    const cellsArray = cells
    // 첫 번째 셀을 기준으로 병합
    const firstCell = cellsArray[0]
    const firstRow = parseInt(firstCell.getAttribute("data-row"))
    const firstCol = parseInt(firstCell.getAttribute("data-col"))

    // 병합할 셀의 범위 계산
    let minRow = firstRow,
      maxRow = firstRow
    let minCol = firstCol,
      maxCol = firstCol

    cellsArray.forEach((cell) => {
      const row = parseInt(cell.getAttribute("data-row"))
      const col = parseInt(cell.getAttribute("data-col"))
      if (row < minRow) minRow = row
      if (row > maxRow) maxRow = row
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    })

    // 병합 범위에 있는 셀들을 숨기고 첫 번째 셀에 rowspan과 colspan 적용
    for (let i = minRow; i <= maxRow; i++) {
      for (let j = minCol; j <= maxCol; j++) {
        const selector = `.cell[data-row='${i}'][data-col='${j}']`
        const cell = gridContainer.querySelector(selector)
        if (cell && cell !== firstCell) {
          cell.style.display = "none"
        }
      }
    }

    // 첫 번째 셀에 rowspan과 colspan 적용
    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1
    firstCell.style.gridRowEnd = `span ${rowspan}`
    firstCell.style.gridColumnEnd = `span ${colspan}`

    // 선택 해제
    clearSelection()
  }

  // 컬럼 리사이징 이벤트 처리
  document.addEventListener("mousemove", (e) => {
    if (isResizing && currentResizer) {
      const dx = e.clientX - startX
      const newWidth = startWidth + dx
      if (newWidth > 50) {
        // 최소 너비 설정
        currentResizer.style.width = `${newWidth}px`
        const colIndex =
          Array.from(currentResizer.parentNode.parentNode.children).indexOf(
            currentResizer.parentNode
          ) - 1 // 첫 셀은 빈 헤더
        const cells = gridContainer.querySelectorAll(
          `.cell[data-col='${colIndex}']`
        )
        cells.forEach((cell) => {
          cell.style.width = `${newWidth}px`
        })
      }
    }

    if (isRowResizing && currentRowResizer) {
      const dy = e.clientY - startY
      const newHeight = startHeight + dy
      if (newHeight > 20) {
        // 최소 높이 설정
        currentRowResizer.style.height = `${newHeight}px`
        const rowIndex = parseInt(currentRowResizer.textContent) - 1
        const cells = gridContainer.querySelectorAll(
          `.cell[data-row='${rowIndex + 1}']`
        )
        cells.forEach((cell) => {
          cell.style.height = `${newHeight}px`
        })
      }
    }
  })

  // 리사이징 종료 이벤트 처리
  document.addEventListener("mouseup", () => {
    if (isResizing) {
      isResizing = false
      currentResizer = null
      document.body.style.cursor = "default"
    }
    if (isRowResizing) {
      isRowResizing = false
      currentRowResizer = null
      document.body.style.cursor = "default"
    }
  })

  // 드래그 앤 드롭 활성화 함수
  function enableDragAndDrop() {
    const cells = gridContainer.querySelectorAll(".cell:not(.header-cell)")
    cells.forEach((cell) => {
      cell.addEventListener("dragstart", (e) => {
        dragStartCell = cell
        e.dataTransfer.setData(
          "text/plain",
          `${cell.getAttribute("data-row")}-${cell.getAttribute("data-col")}`
        )
        cell.classList.add("dragging")
      })

      cell.addEventListener("dragend", () => {
        dragStartCell.classList.remove("dragging")
      })

      cell.addEventListener("dragover", (e) => {
        e.preventDefault()
        cell.classList.add("drag-over")
      })

      cell.addEventListener("dragleave", () => {
        cell.classList.remove("drag-over")
      })

      cell.addEventListener("drop", (e) => {
        e.preventDefault()
        cell.classList.remove("drag-over")
        const data = e.dataTransfer.getData("text/plain")
        const [startRow, startCol] = data.split("-").map(Number)
        const [endRow, endCol] = [
          cell.getAttribute("data-row"),
          cell.getAttribute("data-col"),
        ].map(Number)

        // 셀 데이터 교환
        const startCellKey = `${startRow}-${startCol}`
        const endCellKey = `${endRow}-${endCol}`
        const temp =
          cellData[startCellKey] ||
          `${String.fromCharCode(65 + startCol)}${startRow}`
        cellData[startCellKey] =
          cellData[endCellKey] || `${String.fromCharCode(65 + endCol)}${endRow}`
        cellData[endCellKey] = temp

        // 그리드 재렌더링
        renderGrid()
      })
    })
  }

  // 컨텍스트 메뉴 표시
  gridContainer.addEventListener("contextmenu", (e) => {
    e.preventDefault()
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      contextTargetCell = target
      showContextMenu(e.pageX, e.pageY)
    }
  })

  // 컨텍스트 메뉴 숨기기
  document.addEventListener("click", () => {
    hideContextMenu()
  })

  // 컨텍스트 메뉴 항목 클릭 처리
  contextMenu.addEventListener("click", (e) => {
    const action = e.target.id
    switch (action) {
      case "context-edit":
        enableEditing(contextTargetCell)
        break
      case "context-bg-color":
        changeBackgroundColor([contextTargetCell])
        break
      case "context-text-color":
        changeTextColor([contextTargetCell])
        break
      case "context-text-align-left":
        applyTextAlignToCells("left", [contextTargetCell])
        break
      case "context-text-align-center":
        applyTextAlignToCells("center", [contextTargetCell])
        break
      case "context-text-align-right":
        applyTextAlignToCells("right", [contextTargetCell])
        break
      case "context-merge-cells":
        mergeSelectedCells([contextTargetCell])
        break
      default:
        break
    }
    hideContextMenu()
  })

  // 컨텍스트 메뉴 표시 함수
  function showContextMenu(x, y) {
    contextMenu.style.top = `${y}px`
    contextMenu.style.left = `${x}px`
    contextMenu.style.display = "block"
  }

  // 컨텍스트 메뉴 숨기기 함수
  function hideContextMenu() {
    contextMenu.style.display = "none"
  }

  // 단축키 이벤트 처리
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.metaKey) {
      // Ctrl 또는 Command 키 조합
      switch (e.key.toLowerCase()) {
        case "b": // Ctrl+B: 배경색 변경
          e.preventDefault()
          changeBackgroundColor(Array.from(selectedCells))
          break
        case "c": // Ctrl+C: 글자색 변경
          e.preventDefault()
          changeTextColor(Array.from(selectedCells))
          break
        case "l": // Ctrl+L: 왼쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("left", Array.from(selectedCells))
          break
        case "e": // Ctrl+E: 가운데 정렬
          e.preventDefault()
          applyTextAlignToCells("center", Array.from(selectedCells))
          break
        case "r": // Ctrl+R: 오른쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("right", Array.from(selectedCells))
          break
        default:
          break
      }
    }

    // Esc 키로 선택 해제
    if (e.key === "Escape") {
      clearSelection()
    }
  })

  // 배경색 변경 함수
  function changeBackgroundColor(cells) {
    const color = prompt(
      "원하는 배경색을 입력하세요 (예: #ff0000 또는 red):",
      "#ffff00"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.backgroundColor = color
      })
    }
  }

  // 글자색 변경 함수
  function changeTextColor(cells) {
    const color = prompt(
      "원하는 글자색을 입력하세요 (예: #ff0000 또는 red):",
      "#000000"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.color = color
      })
    }
  }

  // 텍스트 정렬 함수
  function applyTextAlignToCells(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  // 셀 병합 함수
  function mergeSelectedCells(cells = Array.from(selectedCells)) {
    if (cells.length <= 1) {
      alert("병합할 셀을 두 개 이상 선택하세요.")
      return
    }

    // 선택된 셀들을 배열로 변환
    const cellsArray = cells
    // 첫 번째 셀을 기준으로 병합
    const firstCell = cellsArray[0]
    const firstRow = parseInt(firstCell.getAttribute("data-row"))
    const firstCol = parseInt(firstCell.getAttribute("data-col"))

    // 병합할 셀의 범위 계산
    let minRow = firstRow,
      maxRow = firstRow
    let minCol = firstCol,
      maxCol = firstCol

    cellsArray.forEach((cell) => {
      const row = parseInt(cell.getAttribute("data-row"))
      const col = parseInt(cell.getAttribute("data-col"))
      if (row < minRow) minRow = row
      if (row > maxRow) maxRow = row
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    })

    // 병합 범위에 있는 셀들을 숨기고 첫 번째 셀에 rowspan과 colspan 적용
    for (let i = minRow; i <= maxRow; i++) {
      for (let j = minCol; j <= maxCol; j++) {
        const selector = `.cell[data-row='${i}'][data-col='${j}']`
        const cell = gridContainer.querySelector(selector)
        if (cell && cell !== firstCell) {
          cell.style.display = "none"
        }
      }
    }

    // 첫 번째 셀에 rowspan과 colspan 적용
    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1
    firstCell.style.gridRowEnd = `span ${rowspan}`
    firstCell.style.gridColumnEnd = `span ${colspan}`

    // 선택 해제
    clearSelection()
  }

  // 컬럼 리사이징 이벤트 처리
  document.addEventListener("mousemove", (e) => {
    if (isResizing && currentResizer) {
      const dx = e.clientX - startX
      const newWidth = startWidth + dx
      if (newWidth > 50) {
        // 최소 너비 설정
        currentResizer.style.width = `${newWidth}px`
        const colIndex =
          Array.from(currentResizer.parentNode.parentNode.children).indexOf(
            currentResizer.parentNode
          ) - 1 // 첫 셀은 빈 헤더
        const cells = gridContainer.querySelectorAll(
          `.cell[data-col='${colIndex}']`
        )
        cells.forEach((cell) => {
          cell.style.width = `${newWidth}px`
        })
      }
    }

    if (isRowResizing && currentRowResizer) {
      const dy = e.clientY - startY
      const newHeight = startHeight + dy
      if (newHeight > 20) {
        // 최소 높이 설정
        currentRowResizer.style.height = `${newHeight}px`
        const rowIndex = parseInt(currentRowResizer.textContent) - 1
        const cells = gridContainer.querySelectorAll(
          `.cell[data-row='${rowIndex + 1}']`
        )
        cells.forEach((cell) => {
          cell.style.height = `${newHeight}px`
        })
      }
    }
  })

  // 리사이징 종료 이벤트 처리
  document.addEventListener("mouseup", () => {
    if (isResizing) {
      isResizing = false
      currentResizer = null
      document.body.style.cursor = "default"
    }
    if (isRowResizing) {
      isRowResizing = false
      currentRowResizer = null
      document.body.style.cursor = "default"
    }
  })

  // 드래그 앤 드롭 활성화 함수
  function enableDragAndDrop() {
    const cells = gridContainer.querySelectorAll(".cell:not(.header-cell)")
    cells.forEach((cell) => {
      cell.addEventListener("dragstart", (e) => {
        dragStartCell = cell
        e.dataTransfer.setData(
          "text/plain",
          `${cell.getAttribute("data-row")}-${cell.getAttribute("data-col")}`
        )
        cell.classList.add("dragging")
      })

      cell.addEventListener("dragend", () => {
        dragStartCell.classList.remove("dragging")
      })

      cell.addEventListener("dragover", (e) => {
        e.preventDefault()
        cell.classList.add("drag-over")
      })

      cell.addEventListener("dragleave", () => {
        cell.classList.remove("drag-over")
      })

      cell.addEventListener("drop", (e) => {
        e.preventDefault()
        cell.classList.remove("drag-over")
        const data = e.dataTransfer.getData("text/plain")
        const [startRow, startCol] = data.split("-").map(Number)
        const [endRow, endCol] = [
          cell.getAttribute("data-row"),
          cell.getAttribute("data-col"),
        ].map(Number)

        // 셀 데이터 교환
        const startCellKey = `${startRow}-${startCol}`
        const endCellKey = `${endRow}-${endCol}`
        const temp =
          cellData[startCellKey] ||
          `${String.fromCharCode(65 + startCol)}${startRow}`
        cellData[startCellKey] =
          cellData[endCellKey] || `${String.fromCharCode(65 + endCol)}${endRow}`
        cellData[endCellKey] = temp

        // 그리드 재렌더링
        renderGrid()
      })
    })
  }

  // 컨텍스트 메뉴 표시
  gridContainer.addEventListener("contextmenu", (e) => {
    e.preventDefault()
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      contextTargetCell = target
      showContextMenu(e.pageX, e.pageY)
    }
  })

  // 컨텍스트 메뉴 숨기기
  document.addEventListener("click", () => {
    hideContextMenu()
  })

  // 컨텍스트 메뉴 항목 클릭 처리
  contextMenu.addEventListener("click", (e) => {
    const action = e.target.id
    switch (action) {
      case "context-edit":
        enableEditing(contextTargetCell)
        break
      case "context-bg-color":
        changeBackgroundColor([contextTargetCell])
        break
      case "context-text-color":
        changeTextColor([contextTargetCell])
        break
      case "context-text-align-left":
        applyTextAlignToCells("left", [contextTargetCell])
        break
      case "context-text-align-center":
        applyTextAlignToCells("center", [contextTargetCell])
        break
      case "context-text-align-right":
        applyTextAlignToCells("right", [contextTargetCell])
        break
      case "context-merge-cells":
        mergeSelectedCells([contextTargetCell])
        break
      default:
        break
    }
    hideContextMenu()
  })

  // 컨텍스트 메뉴 표시 함수
  function showContextMenu(x, y) {
    contextMenu.style.top = `${y}px`
    contextMenu.style.left = `${x}px`
    contextMenu.style.display = "block"
  }

  // 컨텍스트 메뉴 숨기기 함수
  function hideContextMenu() {
    contextMenu.style.display = "none"
  }

  // 단축키 이벤트 처리
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.metaKey) {
      // Ctrl 또는 Command 키 조합
      switch (e.key.toLowerCase()) {
        case "b": // Ctrl+B: 배경색 변경
          e.preventDefault()
          changeBackgroundColor(Array.from(selectedCells))
          break
        case "c": // Ctrl+C: 글자색 변경
          e.preventDefault()
          changeTextColor(Array.from(selectedCells))
          break
        case "l": // Ctrl+L: 왼쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("left", Array.from(selectedCells))
          break
        case "e": // Ctrl+E: 가운데 정렬
          e.preventDefault()
          applyTextAlignToCells("center", Array.from(selectedCells))
          break
        case "r": // Ctrl+R: 오른쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("right", Array.from(selectedCells))
          break
        default:
          break
      }
    }

    // Esc 키로 선택 해제
    if (e.key === "Escape") {
      clearSelection()
    }
  })

  // 배경색 변경 함수
  function changeBackgroundColor(cells) {
    const color = prompt(
      "원하는 배경색을 입력하세요 (예: #ff0000 또는 red):",
      "#ffff00"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.backgroundColor = color
      })
    }
  }

  // 글자색 변경 함수
  function changeTextColor(cells) {
    const color = prompt(
      "원하는 글자색을 입력하세요 (예: #ff0000 또는 red):",
      "#000000"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.color = color
      })
    }
  }

  // 텍스트 정렬 함수
  function applyTextAlignToCells(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  // 셀 병합 함수
  function mergeSelectedCells(cells = Array.from(selectedCells)) {
    if (cells.length <= 1) {
      alert("병합할 셀을 두 개 이상 선택하세요.")
      return
    }

    // 선택된 셀들을 배열로 변환
    const cellsArray = cells
    // 첫 번째 셀을 기준으로 병합
    const firstCell = cellsArray[0]
    const firstRow = parseInt(firstCell.getAttribute("data-row"))
    const firstCol = parseInt(firstCell.getAttribute("data-col"))

    // 병합할 셀의 범위 계산
    let minRow = firstRow,
      maxRow = firstRow
    let minCol = firstCol,
      maxCol = firstCol

    cellsArray.forEach((cell) => {
      const row = parseInt(cell.getAttribute("data-row"))
      const col = parseInt(cell.getAttribute("data-col"))
      if (row < minRow) minRow = row
      if (row > maxRow) maxRow = row
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    })

    // 병합 범위에 있는 셀들을 숨기고 첫 번째 셀에 rowspan과 colspan 적용
    for (let i = minRow; i <= maxRow; i++) {
      for (let j = minCol; j <= maxCol; j++) {
        const selector = `.cell[data-row='${i}'][data-col='${j}']`
        const cell = gridContainer.querySelector(selector)
        if (cell && cell !== firstCell) {
          cell.style.display = "none"
        }
      }
    }

    // 첫 번째 셀에 rowspan과 colspan 적용
    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1
    firstCell.style.gridRowEnd = `span ${rowspan}`
    firstCell.style.gridColumnEnd = `span ${colspan}`

    // 선택 해제
    clearSelection()
  }

  // 컬럼 리사이징 핸들러 추가
  function addColumnResizers() {
    const headers = gridContainer.querySelectorAll(".header-cell")
    headers.forEach((header) => {
      if (header.classList.contains("row-header")) return // 행 헤더는 제외
      const resizer = document.createElement("div")
      resizer.classList.add("column-resizer")
      header.appendChild(resizer)

      resizer.addEventListener("mousedown", (e) => {
        isResizing = true
        currentResizer = header
        startX = e.clientX
        startWidth = header.offsetWidth
        document.body.style.cursor = "col-resize"
        e.preventDefault()
      })
    })
  }

  // 로우 리사이징 핸들러 추가
  function addRowResizers() {
    const rowHeaders = gridContainer.querySelectorAll(".row-header")
    rowHeaders.forEach((header) => {
      const resizer = document.createElement("div")
      resizer.classList.add("row-resizer")
      header.appendChild(resizer)

      resizer.addEventListener("mousedown", (e) => {
        isRowResizing = true
        currentRowResizer = header
        startY = e.clientY
        startHeight = header.offsetHeight
        document.body.style.cursor = "row-resize"
        e.preventDefault()
      })
    })
  }

  // 드래그 앤 드롭 활성화 함수
  function enableDragAndDrop() {
    const cells = gridContainer.querySelectorAll(".cell:not(.header-cell)")
    cells.forEach((cell) => {
      cell.addEventListener("dragstart", (e) => {
        dragStartCell = cell
        e.dataTransfer.setData(
          "text/plain",
          `${cell.getAttribute("data-row")}-${cell.getAttribute("data-col")}`
        )
        cell.classList.add("dragging")
      })

      cell.addEventListener("dragend", () => {
        dragStartCell.classList.remove("dragging")
      })

      cell.addEventListener("dragover", (e) => {
        e.preventDefault()
        cell.classList.add("drag-over")
      })

      cell.addEventListener("dragleave", () => {
        cell.classList.remove("drag-over")
      })

      cell.addEventListener("drop", (e) => {
        e.preventDefault()
        cell.classList.remove("drag-over")
        const data = e.dataTransfer.getData("text/plain")
        const [startRow, startCol] = data.split("-").map(Number)
        const [endRow, endCol] = [
          cell.getAttribute("data-row"),
          cell.getAttribute("data-col"),
        ].map(Number)

        // 셀 데이터 교환
        const startCellKey = `${startRow}-${startCol}`
        const endCellKey = `${endRow}-${endCol}`
        const temp =
          cellData[startCellKey] ||
          `${String.fromCharCode(65 + startCol)}${startRow}`
        cellData[startCellKey] =
          cellData[endCellKey] || `${String.fromCharCode(65 + endCol)}${endRow}`
        cellData[endCellKey] = temp

        // 그리드 재렌더링
        renderGrid()
      })
    })
  }

  // 컨텍스트 메뉴 표시
  gridContainer.addEventListener("contextmenu", (e) => {
    e.preventDefault()
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      contextTargetCell = target
      showContextMenu(e.pageX, e.pageY)
    }
  })

  // 컨텍스트 메뉴 숨기기
  document.addEventListener("click", () => {
    hideContextMenu()
  })

  // 컨텍스트 메뉴 항목 클릭 처리
  contextMenu.addEventListener("click", (e) => {
    const action = e.target.id
    switch (action) {
      case "context-edit":
        enableEditing(contextTargetCell)
        break
      case "context-bg-color":
        changeBackgroundColor([contextTargetCell])
        break
      case "context-text-color":
        changeTextColor([contextTargetCell])
        break
      case "context-text-align-left":
        applyTextAlignToCells("left", [contextTargetCell])
        break
      case "context-text-align-center":
        applyTextAlignToCells("center", [contextTargetCell])
        break
      case "context-text-align-right":
        applyTextAlignToCells("right", [contextTargetCell])
        break
      case "context-merge-cells":
        mergeSelectedCells([contextTargetCell])
        break
      default:
        break
    }
    hideContextMenu()
  })

  // 컨텍스트 메뉴 표시 함수
  function showContextMenu(x, y) {
    contextMenu.style.top = `${y}px`
    contextMenu.style.left = `${x}px`
    contextMenu.style.display = "block"
  }

  // 컨텍스트 메뉴 숨기기 함수
  function hideContextMenu() {
    contextMenu.style.display = "none"
  }

  // 단축키 이벤트 처리
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.metaKey) {
      // Ctrl 또는 Command 키 조합
      switch (e.key.toLowerCase()) {
        case "b": // Ctrl+B: 배경색 변경
          e.preventDefault()
          changeBackgroundColor(Array.from(selectedCells))
          break
        case "c": // Ctrl+C: 글자색 변경
          e.preventDefault()
          changeTextColor(Array.from(selectedCells))
          break
        case "l": // Ctrl+L: 왼쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("left", Array.from(selectedCells))
          break
        case "e": // Ctrl+E: 가운데 정렬
          e.preventDefault()
          applyTextAlignToCells("center", Array.from(selectedCells))
          break
        case "r": // Ctrl+R: 오른쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("right", Array.from(selectedCells))
          break
        default:
          break
      }
    }

    // Esc 키로 선택 해제
    if (e.key === "Escape") {
      clearSelection()
    }
  })

  // 배경색 변경 함수
  function changeBackgroundColor(cells) {
    const color = prompt(
      "원하는 배경색을 입력하세요 (예: #ff0000 또는 red):",
      "#ffff00"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.backgroundColor = color
      })
    }
  }

  // 글자색 변경 함수
  function changeTextColor(cells) {
    const color = prompt(
      "원하는 글자색을 입력하세요 (예: #ff0000 또는 red):",
      "#000000"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.color = color
      })
    }
  }

  // 텍스트 정렬 함수
  function applyTextAlignToCells(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  // 셀 병합 함수
  function mergeSelectedCells(cells = Array.from(selectedCells)) {
    if (cells.length <= 1) {
      alert("병합할 셀을 두 개 이상 선택하세요.")
      return
    }

    // 선택된 셀들을 배열로 변환
    const cellsArray = cells
    // 첫 번째 셀을 기준으로 병합
    const firstCell = cellsArray[0]
    const firstRow = parseInt(firstCell.getAttribute("data-row"))
    const firstCol = parseInt(firstCell.getAttribute("data-col"))

    // 병합할 셀의 범위 계산
    let minRow = firstRow,
      maxRow = firstRow
    let minCol = firstCol,
      maxCol = firstCol

    cellsArray.forEach((cell) => {
      const row = parseInt(cell.getAttribute("data-row"))
      const col = parseInt(cell.getAttribute("data-col"))
      if (row < minRow) minRow = row
      if (row > maxRow) maxRow = row
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    })

    // 병합 범위에 있는 셀들을 숨기고 첫 번째 셀에 rowspan과 colspan 적용
    for (let i = minRow; i <= maxRow; i++) {
      for (let j = minCol; j <= maxCol; j++) {
        const selector = `.cell[data-row='${i}'][data-col='${j}']`
        const cell = gridContainer.querySelector(selector)
        if (cell && cell !== firstCell) {
          cell.style.display = "none"
        }
      }
    }

    // 첫 번째 셀에 rowspan과 colspan 적용
    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1
    firstCell.style.gridRowEnd = `span ${rowspan}`
    firstCell.style.gridColumnEnd = `span ${colspan}`

    // 선택 해제
    clearSelection()
  }

  // 컬럼 리사이징 이벤트 처리
  document.addEventListener("mousemove", (e) => {
    if (isResizing && currentResizer) {
      const dx = e.clientX - startX
      const newWidth = startWidth + dx
      if (newWidth > 50) {
        // 최소 너비 설정
        currentResizer.style.width = `${newWidth}px`
        const colIndex =
          Array.from(currentResizer.parentNode.parentNode.children).indexOf(
            currentResizer.parentNode
          ) - 1 // 첫 셀은 빈 헤더
        const cells = gridContainer.querySelectorAll(
          `.cell[data-col='${colIndex}']`
        )
        cells.forEach((cell) => {
          cell.style.width = `${newWidth}px`
        })
      }
    }

    if (isRowResizing && currentRowResizer) {
      const dy = e.clientY - startY
      const newHeight = startHeight + dy
      if (newHeight > 20) {
        // 최소 높이 설정
        currentRowResizer.style.height = `${newHeight}px`
        const rowIndex = parseInt(currentRowResizer.textContent) - 1
        const cells = gridContainer.querySelectorAll(
          `.cell[data-row='${rowIndex + 1}']`
        )
        cells.forEach((cell) => {
          cell.style.height = `${newHeight}px`
        })
      }
    }
  })

  // 리사이징 종료 이벤트 처리
  document.addEventListener("mouseup", () => {
    if (isResizing) {
      isResizing = false
      currentResizer = null
      document.body.style.cursor = "default"
    }
    if (isRowResizing) {
      isRowResizing = false
      currentRowResizer = null
      document.body.style.cursor = "default"
    }
  })

  // 드래그 앤 드롭 활성화 함수
  function enableDragAndDrop() {
    const cells = gridContainer.querySelectorAll(".cell:not(.header-cell)")
    cells.forEach((cell) => {
      cell.addEventListener("dragstart", (e) => {
        dragStartCell = cell
        e.dataTransfer.setData(
          "text/plain",
          `${cell.getAttribute("data-row")}-${cell.getAttribute("data-col")}`
        )
        cell.classList.add("dragging")
      })

      cell.addEventListener("dragend", () => {
        dragStartCell.classList.remove("dragging")
      })

      cell.addEventListener("dragover", (e) => {
        e.preventDefault()
        cell.classList.add("drag-over")
      })

      cell.addEventListener("dragleave", () => {
        cell.classList.remove("drag-over")
      })

      cell.addEventListener("drop", (e) => {
        e.preventDefault()
        cell.classList.remove("drag-over")
        const data = e.dataTransfer.getData("text/plain")
        const [startRow, startCol] = data.split("-").map(Number)
        const [endRow, endCol] = [
          cell.getAttribute("data-row"),
          cell.getAttribute("data-col"),
        ].map(Number)

        // 셀 데이터 교환
        const startCellKey = `${startRow}-${startCol}`
        const endCellKey = `${endRow}-${endCol}`
        const temp =
          cellData[startCellKey] ||
          `${String.fromCharCode(65 + startCol)}${startRow}`
        cellData[startCellKey] =
          cellData[endCellKey] || `${String.fromCharCode(65 + endCol)}${endRow}`
        cellData[endCellKey] = temp

        // 그리드 재렌더링
        renderGrid()
      })
    })
  }

  // 컨텍스트 메뉴 표시
  gridContainer.addEventListener("contextmenu", (e) => {
    e.preventDefault()
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      contextTargetCell = target
      showContextMenu(e.pageX, e.pageY)
    }
  })

  // 컨텍스트 메뉴 숨기기
  document.addEventListener("click", () => {
    hideContextMenu()
  })

  // 컨텍스트 메뉴 항목 클릭 처리
  contextMenu.addEventListener("click", (e) => {
    const action = e.target.id
    switch (action) {
      case "context-edit":
        enableEditing(contextTargetCell)
        break
      case "context-bg-color":
        changeBackgroundColor([contextTargetCell])
        break
      case "context-text-color":
        changeTextColor([contextTargetCell])
        break
      case "context-text-align-left":
        applyTextAlignToCells("left", [contextTargetCell])
        break
      case "context-text-align-center":
        applyTextAlignToCells("center", [contextTargetCell])
        break
      case "context-text-align-right":
        applyTextAlignToCells("right", [contextTargetCell])
        break
      case "context-merge-cells":
        mergeSelectedCells([contextTargetCell])
        break
      default:
        break
    }
    hideContextMenu()
  })

  // 컨텍스트 메뉴 표시 함수
  function showContextMenu(x, y) {
    contextMenu.style.top = `${y}px`
    contextMenu.style.left = `${x}px`
    contextMenu.style.display = "block"
  }

  // 컨텍스트 메뉴 숨기기 함수
  function hideContextMenu() {
    contextMenu.style.display = "none"
  }

  // 단축키 이벤트 처리
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.metaKey) {
      // Ctrl 또는 Command 키 조합
      switch (e.key.toLowerCase()) {
        case "b": // Ctrl+B: 배경색 변경
          e.preventDefault()
          changeBackgroundColor(Array.from(selectedCells))
          break
        case "c": // Ctrl+C: 글자색 변경
          e.preventDefault()
          changeTextColor(Array.from(selectedCells))
          break
        case "l": // Ctrl+L: 왼쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("left", Array.from(selectedCells))
          break
        case "e": // Ctrl+E: 가운데 정렬
          e.preventDefault()
          applyTextAlignToCells("center", Array.from(selectedCells))
          break
        case "r": // Ctrl+R: 오른쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("right", Array.from(selectedCells))
          break
        default:
          break
      }
    }

    // Esc 키로 선택 해제
    if (e.key === "Escape") {
      clearSelection()
    }
  })

  // 배경색 변경 함수
  function changeBackgroundColor(cells) {
    const color = prompt(
      "원하는 배경색을 입력하세요 (예: #ff0000 또는 red):",
      "#ffff00"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.backgroundColor = color
      })
    }
  }

  // 글자색 변경 함수
  function changeTextColor(cells) {
    const color = prompt(
      "원하는 글자색을 입력하세요 (예: #ff0000 또는 red):",
      "#000000"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.color = color
      })
    }
  }

  // 텍스트 정렬 함수
  function applyTextAlignToCells(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  // 셀 병합 함수
  function mergeSelectedCells(cells = Array.from(selectedCells)) {
    if (cells.length <= 1) {
      alert("병합할 셀을 두 개 이상 선택하세요.")
      return
    }

    // 선택된 셀들을 배열로 변환
    const cellsArray = cells
    // 첫 번째 셀을 기준으로 병합
    const firstCell = cellsArray[0]
    const firstRow = parseInt(firstCell.getAttribute("data-row"))
    const firstCol = parseInt(firstCell.getAttribute("data-col"))

    // 병합할 셀의 범위 계산
    let minRow = firstRow,
      maxRow = firstRow
    let minCol = firstCol,
      maxCol = firstCol

    cellsArray.forEach((cell) => {
      const row = parseInt(cell.getAttribute("data-row"))
      const col = parseInt(cell.getAttribute("data-col"))
      if (row < minRow) minRow = row
      if (row > maxRow) maxRow = row
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    })

    // 병합 범위에 있는 셀들을 숨기고 첫 번째 셀에 rowspan과 colspan 적용
    for (let i = minRow; i <= maxRow; i++) {
      for (let j = minCol; j <= maxCol; j++) {
        const selector = `.cell[data-row='${i}'][data-col='${j}']`
        const cell = gridContainer.querySelector(selector)
        if (cell && cell !== firstCell) {
          cell.style.display = "none"
        }
      }
    }

    // 첫 번째 셀에 rowspan과 colspan 적용
    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1
    firstCell.style.gridRowEnd = `span ${rowspan}`
    firstCell.style.gridColumnEnd = `span ${colspan}`

    // 선택 해제
    clearSelection()
  }

  // 컬럼 리사이징 핸들러 추가
  function addColumnResizers() {
    const headers = gridContainer.querySelectorAll(".header-cell")
    headers.forEach((header) => {
      if (header.classList.contains("row-header")) return // 행 헤더는 제외
      const resizer = document.createElement("div")
      resizer.classList.add("column-resizer")
      header.appendChild(resizer)

      resizer.addEventListener("mousedown", (e) => {
        isResizing = true
        currentResizer = header
        startX = e.clientX
        startWidth = header.offsetWidth
        document.body.style.cursor = "col-resize"
        e.preventDefault()
      })
    })
  }

  // 로우 리사이징 핸들러 추가
  function addRowResizers() {
    const rowHeaders = gridContainer.querySelectorAll(".row-header")
    rowHeaders.forEach((header) => {
      const resizer = document.createElement("div")
      resizer.classList.add("row-resizer")
      header.appendChild(resizer)

      resizer.addEventListener("mousedown", (e) => {
        isRowResizing = true
        currentRowResizer = header
        startY = e.clientY
        startHeight = header.offsetHeight
        document.body.style.cursor = "row-resize"
        e.preventDefault()
      })
    })
  }

  // 컨텍스트 메뉴 표시 함수
  function showContextMenu(x, y) {
    contextMenu.style.top = `${y}px`
    contextMenu.style.left = `${x}px`
    contextMenu.style.display = "block"
  }

  // 컨텍스트 메뉴 숨기기 함수
  function hideContextMenu() {
    contextMenu.style.display = "none"
  }

  // 단축키 이벤트 처리
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey || e.metaKey) {
      // Ctrl 또는 Command 키 조합
      switch (e.key.toLowerCase()) {
        case "b": // Ctrl+B: 배경색 변경
          e.preventDefault()
          changeBackgroundColor(Array.from(selectedCells))
          break
        case "c": // Ctrl+C: 글자색 변경
          e.preventDefault()
          changeTextColor(Array.from(selectedCells))
          break
        case "l": // Ctrl+L: 왼쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("left", Array.from(selectedCells))
          break
        case "e": // Ctrl+E: 가운데 정렬
          e.preventDefault()
          applyTextAlignToCells("center", Array.from(selectedCells))
          break
        case "r": // Ctrl+R: 오른쪽 정렬
          e.preventDefault()
          applyTextAlignToCells("right", Array.from(selectedCells))
          break
        default:
          break
      }
    }

    // Esc 키로 선택 해제
    if (e.key === "Escape") {
      clearSelection()
    }
  })

  // 배경색 변경 함수
  function changeBackgroundColor(cells) {
    const color = prompt(
      "원하는 배경색을 입력하세요 (예: #ff0000 또는 red):",
      "#ffff00"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.backgroundColor = color
      })
    }
  }

  // 글자색 변경 함수
  function changeTextColor(cells) {
    const color = prompt(
      "원하는 글자색을 입력하세요 (예: #ff0000 또는 red):",
      "#000000"
    )
    if (color) {
      cells.forEach((cell) => {
        cell.style.color = color
      })
    }
  }

  // 텍스트 정렬 함수
  function applyTextAlignToCells(align, cells) {
    cells.forEach((cell) => {
      cell.style.textAlign = align
    })
  }

  // 셀 병합 함수
  function mergeSelectedCells(cells = Array.from(selectedCells)) {
    if (cells.length <= 1) {
      alert("병합할 셀을 두 개 이상 선택하세요.")
      return
    }

    // 선택된 셀들을 배열로 변환
    const cellsArray = cells
    // 첫 번째 셀을 기준으로 병합
    const firstCell = cellsArray[0]
    const firstRow = parseInt(firstCell.getAttribute("data-row"))
    const firstCol = parseInt(firstCell.getAttribute("data-col"))

    // 병합할 셀의 범위 계산
    let minRow = firstRow,
      maxRow = firstRow
    let minCol = firstCol,
      maxCol = firstCol

    cellsArray.forEach((cell) => {
      const row = parseInt(cell.getAttribute("data-row"))
      const col = parseInt(cell.getAttribute("data-col"))
      if (row < minRow) minRow = row
      if (row > maxRow) maxRow = row
      if (col < minCol) minCol = col
      if (col > maxCol) maxCol = col
    })

    // 병합 범위에 있는 셀들을 숨기고 첫 번째 셀에 rowspan과 colspan 적용
    for (let i = minRow; i <= maxRow; i++) {
      for (let j = minCol; j <= maxCol; j++) {
        const selector = `.cell[data-row='${i}'][data-col='${j}']`
        const cell = gridContainer.querySelector(selector)
        if (cell && cell !== firstCell) {
          cell.style.display = "none"
        }
      }
    }

    // 첫 번째 셀에 rowspan과 colspan 적용
    const rowspan = maxRow - minRow + 1
    const colspan = maxCol - minCol + 1
    firstCell.style.gridRowEnd = `span ${rowspan}`
    firstCell.style.gridColumnEnd = `span ${colspan}`

    // 선택 해제
    clearSelection()
  }

  // 셀 선택 및 키보드 네비게이션
  gridContainer.addEventListener("click", (e) => {
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      currentCell = target
      toggleSelection(target)
    }
  })

  // 셀 더블 클릭 시 편집 모드 활성화
  gridContainer.addEventListener("dblclick", (e) => {
    const target = e.target
    if (
      target.classList.contains("cell") &&
      !target.classList.contains("header-cell")
    ) {
      enableEditing(target)
    }
  })

  // 그리드 초기 렌더링
  renderGrid()
})
