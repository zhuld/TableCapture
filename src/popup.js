document.getElementById('extractBtn').addEventListener('click', async () => {
  const status = document.getElementById('status');
  status.textContent = '正在提取...';
  status.className = 'info';

  try {
    // 获取当前活动标签页
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab.id) throw new Error('无法获取标签页');

    // 注入脚本到当前页面，提取成绩表数据
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: extractGradeTableAsJSON,
    });

    // executeScript 返回的结果数组
    const jsonString = results[0]?.result;
    if (!jsonString) {
      status.textContent = '未发现成绩表，请确认页面中包含表格数据。';
      status.className = 'info';
      return;
    }

    // 触发浏览器下载 JSON 文件
    downloadJSON(jsonString, '成绩表_'+new Date().toLocaleString('en-US') +'.json');
    status.textContent = '提取成功，文件已下载。';
    status.className = 'success';
  } catch (err) {
    status.textContent = '提取失败：' + err.message;
    status.className = 'error';
    console.error(err);
  }
});

/**
 * 在页面上下文中执行的函数：识别成绩表并返回 JSON 字符串
 */
function extractGradeTableAsJSON() {
  // ---------- 1. 获取所有表格 ----------
  let tables = document.querySelectorAll('table');
  const iframes = document.querySelectorAll('iframe');
  for (const iframe of iframes) {
    try {
      const iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
      const iframeTables = iframeDoc.querySelectorAll('table');
      tables = [...tables, ...iframeTables];
    } catch (e) {
      // 无法访问跨域 iframe，忽略
    } 
  }

  if (!tables.length) return null;

  // ---------- 2. 定义表头关键词（中英文） ----------
  const gradeKeywords = [
    '学号', '姓名', '学生', '成绩', '分数', '课程', '科目',
    '语文', '数学', '英语', '物理', '化学', '生物', '历史', '地理', '政治',
    '总分', '平均', '排名', '绩点', '等级',
    'id', 'name', 'student', 'grade', 'score', 'subject', 'course','contact','email',
    'total', 'average', 'rank', 'gpa', 'level','描述','Company','公司'
  ];
  const keywordRegex = new RegExp(gradeKeywords.join('|'), 'i');

  // ---------- 3. 遍历表格，寻找最可能的成绩表 ----------
  let bestTable = null;
  let bestScore = 0;

  for (const table of tables) {
    // 寻找表头行：优先 thead 中的 tr，否则取第一个全部由 th 组成的 tr，否则取第一个 tr
    let headerRow = null;
    const thead = table.querySelector('thead');
    if (thead) {
      headerRow = thead.querySelector('tr');
    }
    if (!headerRow) {
      const rows = table.querySelectorAll('tr');
      for (const row of rows) {
        const cells = row.querySelectorAll('th');
        if (cells.length > 0) {
          headerRow = row;
          break;
        }
      }
    }
    if (!headerRow) {
      headerRow = table.querySelector('tr');
    }
    if (!headerRow) continue;

    // 获取表头文本数组
    const headers = [];
    const headerCells = headerRow.querySelectorAll('th, td');
    for (const cell of headerCells) {
      headers.push(cell.textContent.trim().replace(/\s+/g, ' '));
    }

    // 计算有多少个表头匹配成绩关键词
    let matchCount = 0;
    for (const h of headers) {
      if (keywordRegex.test(h)) matchCount++;
    }

    // 至少匹配2个关键词才认为是成绩表
    if (matchCount >= 2 && matchCount > bestScore) {
      bestScore = matchCount;
      bestTable = table;
    }
  }

  if (!bestTable) return null;

  // ---------- 4. 提取数据行 ----------
  const dataRows = bestTable.querySelectorAll('tr');
  // 重新获取最佳表格的表头（与上相同逻辑，确保稳定）
  let headerRow = bestTable.querySelector('thead tr');
  if (!headerRow) {
    const rows = bestTable.querySelectorAll('tr');
    for (const row of rows) {
      if (row.querySelectorAll('th').length > 0) {
        headerRow = row;
        break;
      }
    }
  }
  if (!headerRow) headerRow = bestTable.querySelector('tr');
  if (!headerRow) return null;

  const headerCells = headerRow.querySelectorAll('th, td');
  const headers = Array.from(headerCells).map(cell => cell.textContent.trim().replace(/\s+/g, ' '));

  // 收集数据（跳过表头行本身）
  const jsonData = [];
  let headerFound = false;
  for (const row of dataRows) {
    if (row === headerRow) {
      headerFound = true;
      continue; // 跳过表头行
    }
    if (!headerFound) continue; // 表头之前的行忽略

    const cells = row.querySelectorAll('td, th');
    if (cells.length !== headers.length) continue; // 列数不一致的跳过（可能是合并行）

    const rowObject = {};
    for (let i = 0; i < headers.length; i++) {
      const key = headers[i] || `列${i + 1}`;
      const value = cells[i].textContent.trim();
      rowObject[key] = value;
    }
    // 可选：过滤掉完全空白的行
    if (Object.values(rowObject).some(v => v !== '')) {
      jsonData.push(rowObject);
    }
  }

  if (jsonData.length === 0) return null;

  // ---------- 5. 返回 JSON 字符串 ----------
  return JSON.stringify(jsonData, null, 2);
}

/**
 * 在浏览器中创建一个下载任务
 * @param {string} jsonString - JSON 字符串
 * @param {string} filename - 文件名
 */
function downloadJSON(jsonString, filename) {
  const blob = new Blob([jsonString], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}