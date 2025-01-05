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
  // 셀 클릭 시 편집 모드로 전환 (기본적으로 contenteditable="true" 설정)
  gridBody.addEventListener('dblclick', (e) => {
    if (e.target.tagName === 'TD') {
      e.target.contentEditable = true;
      e.target.focus();
    }
  });

  gridBody.addEventListener('click', (e) => {
    if (e.target.tagName === 'TD') {
      setActiveCell(e.target);
    }
  });

  // 오잉 왜 키보드는 document에 붙이지??
  document.addEventListener('keydown', (e) => {
    // overflow 시 스크롤 방지?
    e.preventDefault();
    if (!activeCell) return;

    let cellRow = activeCell.dataset.row;
    let cellCol = activeCell.dataset.col;

    // 예외 처리
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

    const element = document.querySelector(
      `td[data-row="${cellRow}"][data-col="${cellCol}"]`
    );

    if (element) {
      setActiveCell(element);

      const worksheet = document.getElementById('worksheet');
      const grid = document.getElementById('grid');
      const stickyWidth = grid.querySelector('th').offsetWidth;
      const cellLeft = element.offsetLeft;
      if (cellLeft < worksheet.scrollLeft) {
        worksheet.scrollLeft = cellLeft - stickyWidth;
      }

      element.scrollIntoView({
        behavior: 'instant',
        block: 'end',
      });
    }
  });
}
