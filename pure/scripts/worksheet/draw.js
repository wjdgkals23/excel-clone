export const INITIAL_GRID_ROW = 100;
export const INITIAL_GRID_COL = 26;

export function layoutSheet() {
  let activeCell = null;
  const girdDiv = document.createElement('div');
  girdDiv.id = 'worksheet';

  const table = document.createElement('table');
  girdDiv.appendChild(table);

  table.id = 'grid';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.append(thead);
  table.append(tbody);

  document.body.appendChild(girdDiv);

  return (function (_activeCell) {
    drawSheet();
    addEvent(_activeCell);
    addMouseEvent();
  })(activeCell);
}

function drawSheet() {
  const gridBody = document
    .getElementById('grid')
    .getElementsByTagName('tbody')[0];
  const gridHeader = document
    .getElementById('grid')
    .getElementsByTagName('thead')[0];

  // 초기 행과 열 생성
  const createGrid = (rows, cols) => {
    const headerTr = document.createElement('tr');
    const originPoint = document.createElement('th');
    originPoint.id = 'origin-point';
    headerTr.appendChild(originPoint);
    for (let i = 0; i < cols; i++) {
      const cell = document.createElement('th');
      cell.innerText = String.fromCharCode(65 + i);
      headerTr.appendChild(cell);
    }
    gridHeader.appendChild(headerTr);

    const fragment = document.createDocumentFragment();
    for (let i = 0; i < rows; i++) {
      const row = document.createElement('tr');
      const rowHeader = document.createElement('th');
      rowHeader.textContent = i;
      row.appendChild(rowHeader);

      for (let j = 0; j < cols; j++) {
        const cell = document.createElement('td');
        // cell.contentEditable = 'true';
        cell.setAttribute('data-row', i);
        cell.setAttribute('data-col', j);
        row.appendChild(cell);
      }

      fragment.appendChild(row);
    }
    // 여러번 돔 객체에 직접 접근해서 다수의 노드를 추가해야할 경우, fragment를 퍼포먼스 좋게 처리할 수 있다.
    gridBody.appendChild(fragment);
  };

  createGrid(100, 26); // 100행, 26열 (A-Z)
}

function addEvent(activeCell) {
  const setActiveCell = (cell) => {
    if (activeCell) {
      activeCell.classList.remove('active');
    }
    activeCell = cell;
    activeCell.classList.add('active');
  };

  const gridBody = document
    .getElementById('grid')
    .getElementsByTagName('tbody')[0];
  const gridHeader = document
    .getElementById('grid')
    .getElementsByTagName('thead')[0];
  // 셀 클릭 시 편집 모드로 전환 (기본적으로 contenteditable="true" 설정)
  gridBody.addEventListener('dblclick', (e) => {
    if (e.target.tagName === 'TD') {
      e.target.contentEditable = true;
      e.target.focus();
    }
  });

  gridBody.addEventListener('mousedown', (e) => {
    if (e.target.tagName === 'TD') {
      setActiveCell(e.target);
    }
  });

  document.addEventListener('keydown', (e) => {
    if (!activeCell) return;

    let cellRow = activeCell.dataset.row;
    let cellCol = activeCell.dataset.col;
    let withCtrl = e.ctrlKey || e.metaKey;

    if (
      e.key === 'ArrowDown' ||
      e.key === 'ArrowUp' ||
      e.key === 'ArrowLeft' ||
      e.key === 'ArrowRight'
    )
      e.preventDefault();

    if (withCtrl) {
      switch (e.key) {
        case 'ArrowUp':
          cellRow = 0;
          break;
        case 'ArrowRight': {
          const maxCol = document.querySelector('thead').querySelector('tr')
            .children.length;
          cellCol = maxCol - 2;
          break;
        }
        case 'ArrowDown': {
          const maxRow = document.querySelector('tbody').children.length;
          cellRow = maxRow - 1;
          break;
        }
        case 'ArrowLeft':
          cellCol = 0;
          break;
      }
    } else {
      switch (e.key) {
        case 'ArrowUp':
          cellRow--;
          break;
        case 'ArrowRight':
          cellCol++;
          break;
        case 'ArrowDown':
          cellRow++;
          break;
        case 'ArrowLeft':
          cellCol--;
          break;
      }
    }

    const element = document.querySelector(
      `td[data-row="${cellRow}"][data-col="${cellCol}"]`
    );

    if (element) {
      setActiveCell(element);

      const worksheet = document.getElementById('worksheet');
      const stickyWidth = gridBody.querySelector('th').offsetWidth;
      const stickyHeight = gridHeader.offsetHeight;

      // 셀의 위치와 스크롤 영역 계산
      const cellLeft = element.offsetLeft;
      const cellRight = cellLeft + element.offsetWidth;
      const cellTop = element.offsetTop;
      const cellBottom = cellTop + element.offsetHeight;

      const scrollLeft = worksheet.scrollLeft;
      const scrollTop = worksheet.scrollTop;
      const viewWidth = worksheet.clientWidth;
      const viewHeight = worksheet.clientHeight;

      // 가로 스크롤 조정
      if (cellLeft < scrollLeft + stickyWidth) {
        worksheet.scrollLeft = cellLeft - stickyWidth;
      } else if (cellRight > scrollLeft + viewWidth) {
        worksheet.scrollLeft = cellRight - viewWidth;
      }

      // 세로 스크롤 조정
      if (cellTop < scrollTop + stickyHeight) {
        worksheet.scrollTop = cellTop - stickyHeight;
      } else if (cellBottom > scrollTop + viewHeight) {
        worksheet.scrollTop = cellBottom - viewHeight;
      }
    }
  });
}

function addMouseEvent() {
  const gridBody = document
    .getElementById('grid')
    .getElementsByTagName('tbody')[0];
  let startCell = null;
  let endCell = null;

  const clearSelection = () => {
    // 기존 선택된 셀 초기화
    const selectedCells = gridBody.querySelectorAll('.selected');
    selectedCells.forEach((cell) => cell.classList.remove('selected'));
  };

  const selectCells = (start, end) => {
    clearSelection();

    const startRow = Math.min(start.dataset.row, end.dataset.row);
    const endRow = Math.max(start.dataset.row, end.dataset.row);
    const startCol = Math.min(start.dataset.col, end.dataset.col);
    const endCol = Math.max(start.dataset.col, end.dataset.col);

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = document.querySelector(
          `td[data-row="${row}"][data-col="${col}"]`
        );
        if (cell) {
          cell.classList.add('selected'); // 선택 상태 표시
        }
      }
    }
  };

  gridBody.addEventListener('mousedown', (e) => {
    if (e.target.tagName === 'TD') {
      startCell = e.target;
      endCell = e.target;
      clearSelection();
    }
  });

  gridBody.addEventListener('mousemove', (e) => {
    if (e.target.tagName === 'TD') {
      if ((startCell !== null, endCell !== null)) {
        if (
          JSON.stringify(endCell.dataset) !== JSON.stringify(e.target.dataset)
        ) {
          endCell = e.target;
          selectCells(startCell, endCell);
        }
      }
    }
  });

  gridBody.addEventListener('mouseup', (e) => {
    if (e.target.tagName === 'TD') {
      startCell = null;
      endCell = null;
    }
  });
}
