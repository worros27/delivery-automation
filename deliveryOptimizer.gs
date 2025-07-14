function GETCOMPANYADDRESS(name) {
  if (!name) return '';
  const cache = CacheService.getScriptCache();
  const cacheKey = 'addr_' + name;
  const cached = cache.get(cacheKey);
  if (cached) return cached;

  const API_KEY = 'AIzaSyCKCeG9dezl6fnhUlCuS6MIW2pwvjC45ys';
  const query = name + ', Los Angeles';

  let url = [
    'https://maps.googleapis.com/maps/api/place/findplacefromtext/json',
    '?input=', encodeURIComponent(query),
    '&inputtype=textquery',
    '&fields=formatted_address',
    '&key=', API_KEY
  ].join('');

  let resp = UrlFetchApp.fetch(url);
  let data = JSON.parse(resp.getContentText());
  let addr = (data.candidates && data.candidates[0])
    ? data.candidates[0].formatted_address
    : '';

  if (!addr) {
    url = [
      'https://maps.googleapis.com/maps/api/geocode/json',
      '?address=', encodeURIComponent(query),
      '&key=', API_KEY
    ].join('');
    resp = UrlFetchApp.fetch(url);
    data = JSON.parse(resp.getContentText());
    addr = (data.results && data.results[0])
      ? data.results[0].formatted_address
      : '';
  }

  if (addr) cache.put(cacheKey, addr, 6 * 60 * 60);
  return addr;
}

function autoFillDeliveryFields() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  sheet.getRange("I2").setFormula("=IFERROR(TEXT(INDEX(Sheet1!C:C, MATCH(B2, Sheet1!A:A, 0)), \"HH:MM\"), \"\")");
  sheet.getRange("J2").setFormula("=IFERROR(TEXT(INDEX(Sheet1!D:D, MATCH(B2, Sheet1!A:A, 0)), \"HH:MM\"), \"\")");
  sheet.getRange("K2").setFormula("=IFERROR(INDEX(Sheet1!E:E, MATCH(B2, Sheet1!A:A, 0)), \"\")");

  sheet.getRange("I2:K2").autoFill(sheet.getRange("I2:K" + lastRow), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}
// ─────────────────────────────────────────
// CONFIGURATION – 修改为你的配置 / Change these to your own
// ─────────────────────────────────────────
const API_KEY       = 'AIzaSyCKCeG9dezl6fnhUlCuS6MIW2pwvjC45ys';  // 你的 Google Maps API Key，用于调用 Geocoding 和 Directions API
const WAREHOUSE     = '3171 E 12th St, Los Angeles, CA 90023';    // 仓库地址，所有路线从此出发并回到此地
const ADDRESS_RANGE = 'C2:C';                                   // 供应商地址范围，A1 形式，从 C2 一直到该列最后
const OUTPUT_COL    = 9;                                        // 将访问顺序写入第 9 列 (I 列)
const BATCH_SIZE    = 20;                                       // 每批最多多少个停靠点，不含起终点
const BATCH_COLORS  = [                                         // 每批对应一种底色，循环使用
  '#FFB3BA','#FFDFBA','#FFFFBA','#BAFFC9','#BAE1FF',
  '#E6E6FA','#FFF0F5','#F0FFF0','#F0FFFF','#FFE4E1',
  '#F5DEB3','#E0FFFF','#FAFAD2','#FFDAB9','#D8BFD8',
  '#DDA0DD','#EE82EE','#F0E68C','#ADD8E6','#B0E0E6'
];

/**
 * geocodeAddress(addr)
 * --------------------
 * 将任意文字地址通过 Geocoding API 转换成 "lat,lng" 字符串。
 * 如果服务返回错误或无结果，会抛出异常并提示具体错误码，便于排查地址格式问题。
 */
function geocodeAddress(addr) {
  const url = [
    'https://maps.googleapis.com/maps/api/geocode/json',
    '?address=', encodeURIComponent(addr), // 对地址做 URL 编码，避免特殊字符出错
    '&key=',     API_KEY
  ].join('');
  const resp = UrlFetchApp.fetch(url);
  const js   = JSON.parse(resp.getContentText());
  if (!js.results || !js.results.length) {
    // 抛出带有状态码的错误，方便根据 js.status 判断是 OVER_QUERY_LIMIT、ZERO_RESULTS 等问题
    throw new Error(Geocode failed for "${addr}": ${js.status});
  }
  const loc = js.results[0].geometry.location;
  return ${loc.lat},${loc.lng}; // 返回 "纬度,经度"
}

/**
 * getOptimizedBatchByCoords(startCoord, stopsCoords)
 * --------------------------------------------------
 * 对一批经纬度 (stopsCoords) 调用 Directions API，要求 optimize:true，
 * origin 和 destination 都是 startCoord（仓库坐标），
 * 返回按最优顺序排列的新经纬度数组。
 * 抛出的错误会包含 API 返回的状态，方便排查配额、请求格式等问题。
 */
function getOptimizedBatchByCoords(startCoord, stopsCoords) {
  if (!stopsCoords.length) return []; // 如果没有中途停点，则直接返回空数组

  const base = 'https://maps.googleapis.com/maps/api/directions/json';
  const waypointString = 'optimize:true|' + stopsCoords.join('|');
  const params = {
    origin:      startCoord,
    destination: startCoord,
    waypoints:   waypointString,
    key:         API_KEY
  };
  // 拼接查询字符串
  const url = base + '?' + Object.entries(params)
      .map(([k, v]) => ${k}=${encodeURIComponent(v)})
      .join('&');

  const resp = UrlFetchApp.fetch(url);
  const data = JSON.parse(resp.getContentText());
  if (!data.routes || !data.routes[0]) {
    // 报错时显示返回的 API 状态，帮助判断是否因为超配额或请求参数错误
    throw new Error(Directions API error: ${data.status});
  }
  // data.routes[0].waypoint_order 是一个数组，指示 stopsCoords 的新顺序
  return data.routes[0].waypoint_order.map(i => stopsCoords[i]);
}

/**
 * optimizeRouteWithGeocode()
 * --------------------------
 * 主流程函数：
 * 1) 清空旧的 I 列顺序和表格样式
 * 2) 读取 ADDRESS_RANGE 范围内所有非空地址并记下行号
 * 3) Geocode 仓库地址
 * 4) Geocode 每个供应商地址
 * 5) 分批调用 Directions API 做批次优化
 *    - 每批都从仓库出发，途经该批所有停点，再返回仓库
 *    - 每批内部序号从 1 开始重置
 *    - 按行号回写顺序，并给整行上不同底色区分批次
 * 6) 弹窗提示完成状态
 *
 * 出错场景及含义：
 * - “No addresses found”：ADDRESS_RANGE 里没有有效地址，可能范围写错
 * - “Geocode failed”：某地址转换失败，检查地址拼写或 API 配额
 * - “Directions API error”：Directions 请求失败，查看 data.status 了解详情
 */
function optimizeRouteWithGeocode() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  // 1) 清空旧的访问顺序和行背景色
  if (lastRow >= 2) {
    sheet.getRange(2, OUTPUT_COL, lastRow - 1).clearContent();
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).setBackground(null);
  }

  // 2) 读取所有地址并保存对应表格行号
  const values = sheet.getRange(ADDRESS_RANGE).getValues();
  const rows = [];
  values.forEach((r, i) => {
    const addr = r[0].toString().trim();
    if (addr) {
      // i=0 对应 C2，所以行号是 i+2
      rows.push({ addr: addr, row: i + 2 });
    }
  });
  if (!rows.length) {
    SpreadsheetApp.getUi().alert(⚠️ No addresses found in ${ADDRESS_RANGE});
    return;
  }

  // 3) Geocode 仓库地址
  let warehouseCoord;
  try {
    warehouseCoord = geocodeAddress(WAREHOUSE);
  } catch (e) {
    SpreadsheetApp.getUi().alert(❌ Warehouse geocode failed: ${e.message});
    return;
  }

  // 4) Geocode 每个供应商地址
  let coords;
  try {
    coords = rows.map(o => geocodeAddress(o.addr));
  } catch (e) {
    // geocodeAddress 已弹出具体哪个地址失败
    return;
  }

  // 5) 分批优化并写回
  const numBatches = Math.ceil(rows.length / BATCH_SIZE);
  for (let b = 0; b < numBatches; b++) {
    const startIdx    = b * BATCH_SIZE;
    const sliceRows   = rows.slice(startIdx, startIdx + BATCH_SIZE);
    const sliceCoords = coords.slice(startIdx, startIdx + BATCH_SIZE);
    const color       = BATCH_COLORS[b % BATCH_COLORS.length];

    // 调用 Directions API 做批次优化
    let optimizedCoords;
    try {
      optimizedCoords = getOptimizedBatchByCoords(warehouseCoord, sliceCoords);
    } catch (e) {
      console.warn(Batch ${b+1} optimization failed: ${e.message});
      optimizedCoords = sliceCoords.slice(); // 回退到原始顺序
    }

    // 写回并上色
    optimizedCoords.forEach((c, idxInBatch) => {
      const idx   = sliceCoords.indexOf(c);
      const entry = sliceRows[idx];
      // 写序号 (1-based within batch)
      sheet.getRange(entry.row, OUTPUT_COL).setValue(idxInBatch + 1);
      // 给整行设置背景色，便于区分不同批次
      sheet.getRange(entry.row, 1, 1, sheet.getLastColumn()).setBackground(color);
    });
  }

  // 6) 完成提示
  SpreadsheetApp.getUi().alert(
    ✅ Optimized ${rows.length} stops into ${numBatches} batch(es).
  );
} 这是最短路径的code
