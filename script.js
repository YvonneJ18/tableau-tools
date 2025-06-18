const pptSizes = {
    standard: { w: 960, h: 720, scales: [1, 1.25, 1.5, 1.75, 2] },
    widescreen: { w: 1600, h: 900, scales: [0.75, 1, 1.25, 1.5] }
  };

  let selectedPpt = 'standard';
  let selectedScale = 1;

  const pptGrp = document.getElementById('pptSizeGroup');
  const tableGrp = document.getElementById('tableauScaleGroup');
  const canvasContainer = document.getElementById('rightCanvasContainer');
  const canvas = document.getElementById('previewCanvas');
  const btnCalc = document.getElementById('calcBtn');

  // 输入框元素
  const posX = document.getElementById('posX');
  const posY = document.getElementById('posY');
  const widthInput = document.getElementById('width');
  const heightInput = document.getElementById('height');
  const marginTop = document.getElementById('marginTop');
  const marginLeft = document.getElementById('marginLeft');
  const marginBottom = document.getElementById('marginBottom');
  const marginRight = document.getElementById('marginRight');

  // 结果显示
  const resultX = document.getElementById('resultX');
  const resultY = document.getElementById('resultY');
  const resultWidth = document.getElementById('resultWidth');
  const resultHeight = document.getElementById('resultHeight');

  // 初始化Tableau缩放按钮（展示，但右侧画布不跟随缩放变）
  function initTableButtons() {
    tableGrp.innerHTML = '';
    pptSizes[selectedPpt].scales.forEach(s => {
      const b = document.createElement('button');
      const size = pptSizes[selectedPpt];
      b.textContent = `${s}x (${Math.round(size.w * s)}×${Math.round(size.h * s)})`;
      b.dataset.scale = s;
      if (s === 1) b.classList.add('selected');
      b.onclick = () => {
        tableGrp.querySelectorAll('button').forEach(x => x.classList.remove('selected'));
        b.classList.add('selected');
        selectedScale = s;
      };
      tableGrp.append(b);
    });
    fixButtons('tableauScaleGroup');
  }
  // 统一按钮宽度
  function fixButtons(id) {
    const grp = document.getElementById(id);
    let mx = 0;
    grp.querySelectorAll('button').forEach(b => {
      b.style.width = 'auto';
      mx = Math.max(mx, b.offsetWidth);
    });
    grp.querySelectorAll('button').forEach(b => b.style.width = mx + 'px');
  }

  // PowerPoint页面尺寸按钮事件
  pptGrp.querySelectorAll('button').forEach(b => {
    b.onclick = () => {
      pptGrp.querySelectorAll('button').forEach(x => x.classList.remove('selected'));
      b.classList.add('selected');
      selectedPpt = b.dataset.size;
      selectedScale = 1;
      initTableButtons();
      resizeCanvasAndRedraw();
    };
  });

  // 右侧画布根据PowerPoint页面尺寸和0.75等比缩放调整
  function resizeCanvasAndRedraw() {
    const base = pptSizes[selectedPpt];
    const scaleForCanvas = 1;
    let W = base.w * scaleForCanvas;
    let H = base.h * scaleForCanvas;

    // 限制canvas最大宽高，防止超出容器
    const maxW = window.innerWidth - 420;
    const maxH = window.innerHeight;
    if (W > maxW) {
      H = H * (maxW / W);
      W = maxW;
    }
    if (H > maxH) {
      W = W * (maxH / H);
      H = maxH;
    }

    canvas.width = W;
    canvas.height = H;
    canvas.style.width = Math.round(W) + 'px';
    canvas.style.height = Math.round(H) + 'px';

    drawCanvas();
  }

  // 画布绘制TextBox
  function drawCanvas() {
    const ctx = canvas.getContext('2d');
    const W = canvas.width, H = canvas.height;
    ctx.clearRect(0, 0, W, H);

    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, W, H);

    ctx.strokeStyle = '#888';
    ctx.lineWidth = 2;
    ctx.strokeRect(0, 0, W, H);

    ctx.font = '16px Arial';
    ctx.fillStyle = '#555';
    ctx.fillText(selectedPpt === 'standard' ? 'Standard (4:3)' : 'Widescreen (16:9)', 10, 24);

    // 计算TextBox坐标（inch转px，乘以0.75）
    const inch = 96;
    const pxScale = 0.75;
    const x = parseFloat(posX.value), y = parseFloat(posY.value),
      w = parseFloat(widthInput.value), h = parseFloat(heightInput.value);
    const mt = parseFloat(marginTop.value), ml = parseFloat(marginLeft.value),
      mb = parseFloat(marginBottom.value), mr = parseFloat(marginRight.value);

    const rectX = (x + ml) * inch * pxScale;
    const rectY = (y + mt) * inch * pxScale;
    const rectW = (w - ml - mr) * inch * pxScale;
    const rectH = (h - mt - mb) * inch * pxScale;

    ctx.fillStyle = 'rgba(5,98,138,0.2)';
    ctx.strokeStyle = 'rgba(5,98,138,0.8)';
    ctx.lineWidth = 3;
    ctx.fillRect(rectX, rectY, rectW, rectH);
    ctx.strokeRect(rectX, rectY, rectW, rectH);

    // 绘制文本
    ctx.fillStyle = '#05628A';
    ctx.font = '14px Arial';
    ctx.fillText('Text', rectX + 5, rectY + 20);
  }

  // 计算并显示Container坐标(px)
  function calculate() {
  const inch = 96;
  // 这里必须用当前全局的selectedScale变量
  const pxScale = selectedScale;

  const x = parseFloat(posX.value);
  const y = parseFloat(posY.value);
  const w = parseFloat(width.value);
  const h = parseFloat(height.value);

  const mt = parseFloat(marginTop.value);
  const ml = parseFloat(marginLeft.value);
  const mb = parseFloat(marginBottom.value);
  const mr = parseFloat(marginRight.value);

  // 用视图缩放倍数计算container坐标
  const resX = Math.round((x + ml) * inch * pxScale);
  const resY = Math.round((y + mt) * inch * pxScale);
  const resW = Math.round((w - ml - mr) * inch * pxScale);
  const resH = Math.round((h - mt - mb) * inch * pxScale);

  // 更新界面显示
  resultX.textContent = resX;
  resultY.textContent = resY;
  resultWidth.textContent = resW;
  resultHeight.textContent = resH;
}

  btnCalc.onclick = calculate;

  // 允许通过拖拽右侧TextBox调整位置和尺寸，同时同步更新左侧输入框
  let dragging = false;
  let dragType = null; // "move", "resize-right", "resize-bottom", "resize-corner"
  let dragStart = {};
  let boxStart = {};

  function pointInRect(px, py, rx, ry, rw, rh) {
    return px >= rx && px <= rx + rw && py >= ry && py <= ry + rh;
  }

  function isInResizeZone(px, py, rx, ry, rw, rh) {
    const edgeSize = 10;
    return {
      right: (px >= rx + rw - edgeSize && px <= rx + rw) && (py >= ry && py <= ry + rh),
      bottom: (py >= ry + rh - edgeSize && py <= ry + rh) && (px >= rx && px <= rx + rw),
      corner: (px >= rx + rw - edgeSize && px <= rx + rw) && (py >= ry + rh - edgeSize && py <= ry + rh)
    };
  }

  canvas.addEventListener('mousedown', e => {
    const rect = canvas.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;

    const inch = 96;
    const pxScale = 0.75;
    const x = parseFloat(posX.value), y = parseFloat(posY.value),
      w = parseFloat(widthInput.value), h = parseFloat(heightInput.value);
    const mt = parseFloat(marginTop.value), ml = parseFloat(marginLeft.value),
      mb = parseFloat(marginBottom.value), mr = parseFloat(marginRight.value);

    const boxX = (x + ml) * inch * pxScale;
    const boxY = (y + mt) * inch * pxScale;
    const boxW = (w - ml - mr) * inch * pxScale;
    const boxH = (h - mt - mb) * inch * pxScale;

    if (!pointInRect(mx, my, boxX, boxY, boxW, boxH)) return;

    const resizeZones = isInResizeZone(mx, my, boxX, boxY, boxW, boxH);

    if (resizeZones.corner) dragType = 'resize-corner';
    else if (resizeZones.right) dragType = 'resize-right';
    else if (resizeZones.bottom) dragType = 'resize-bottom';
    else dragType = 'move';

    dragging = true;
    dragStart = { mx, my };
    boxStart = { x, y, w, h };
    e.preventDefault();
  });

  window.addEventListener('mousemove', e => {
    if (!dragging) return;
    const rect = canvas.getBoundingClientRect();
    const mx = e.clientX - rect.left;
    const my = e.clientY - rect.top;
    const dx = mx - dragStart.mx;
    const dy = my - dragStart.my;

    const inch = 96;
    const pxScale = 0.75;
    let newX = boxStart.x;
    let newY = boxStart.y;
    let newW = boxStart.w;
    let newH = boxStart.h;

    // 把像素转换成inch调整
    const dxInch = dx / (inch * pxScale);
    const dyInch = dy / (inch * pxScale);

    if (dragType === 'move') {
      newX = boxStart.x + dxInch;
      newY = boxStart.y + dyInch;
      if (newX < 0) newX = 0;
      if (newY < 0) newY = 0;
    } else if (dragType === 'resize-right') {
      newW = boxStart.w + dxInch;
      if (newW < 0.1) newW = 0.1;
    } else if (dragType === 'resize-bottom') {
      newH = boxStart.h + dyInch;
      if (newH < 0.1) newH = 0.1;
    } else if (dragType === 'resize-corner') {
      newW = boxStart.w + dxInch;
      newH = boxStart.h + dyInch;
      if (newW < 0.1) newW = 0.1;
      if (newH < 0.1) newH = 0.1;
    }

    posX.value = newX.toFixed(2);
    posY.value = newY.toFixed(2);
    widthInput.value = newW.toFixed(2);
    heightInput.value = newH.toFixed(2);

    drawCanvas();
    e.preventDefault();
  });

  window.addEventListener('mouseup', e => {
    dragging = false;
    dragType = null;
  });

  // 当输入框变化时自动重新绘制
  [posX, posY, widthInput, heightInput, marginTop, marginLeft, marginBottom, marginRight].forEach(inp => {
    inp.addEventListener('input', () => {
      drawCanvas();
    });
  });

  // 初始化
  initTableButtons();
  resizeCanvasAndRedraw();
  calculate();

  // 窗口大小变更时重新设置画布尺寸
  window.addEventListener('resize', resizeCanvasAndRedraw);