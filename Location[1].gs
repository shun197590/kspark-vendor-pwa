/**
 * ==========================================================
 * 高雄軟體園區廠商查詢系統
 * Location.gs
 * Version 7.0 - 第27步
 *
 * 功能：
 * 1. 將「進駐位置圖」轉為「位置索引」
 * 2. 建立「位置別名」
 * 3. 自動比對位置圖簡稱與正式廠商名稱
 * 4. 建立「位置配對結果」
 * 5. 將棟別、樓層、位置編號加入 PWA 廠商資料
 * 6. 產生「進駐位置圖」原始儲存格精準連結（第27步）
 * ==========================================================
 */


/**
 * ==========================================================
 * 第21步：將「進駐位置圖」轉成「位置索引」
 * ==========================================================
 */
function buildLocationIndex() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sourceSheet =
    ss.getSheetByName(
      '進駐位置圖'
    );


  if (!sourceSheet) {

    throw new Error(
      '找不到工作表：進駐位置圖'
    );

  }


  let indexSheet =
    ss.getSheetByName(
      '位置索引'
    );


  if (!indexSheet) {

    indexSheet =
      ss.insertSheet(
        '位置索引'
      );

  }


  indexSheet.clear();


  const headers = [

    '棟別',
    '樓層',
    '位置編號',
    '位置圖內容',
    '原始儲存格'

  ];


  const result = [
    headers
  ];


  // 國城C棟
  parseLocationBlock(
    sourceSheet,
    result,
    '國城C棟',
    5,
    13,
    1,
    2,
    9,
    1
  );


  // 新創大樓
  parseLocationBlock(
    sourceSheet,
    result,
    '新創大樓',
    3,
    12,
    11,
    12,
    15,
    1
  );


  // 鴻海大樓
  parseLocationBlock(
    sourceSheet,
    result,
    '鴻海大樓',
    2,
    16,
    18,
    19,
    28,
    1
  );


  // 國城A棟
  parseLocationBlock(
    sourceSheet,
    result,
    '國城A棟',
    18,
    29,
    1,
    2,
    19,
    17
  );


  // 國城B棟
  parseLocationBlock(
    sourceSheet,
    result,
    '國城B棟',
    18,
    29,
    21,
    22,
    25,
    17
  );


  if (result.length > 1) {

    indexSheet
      .getRange(
        1,
        1,
        result.length,
        headers.length
      )
      .setValues(
        result
      );

  }


  indexSheet
    .setFrozenRows(1);


  indexSheet
    .getRange('A1:E1')
    .setFontWeight('bold')
    .setBackground('#075d97')
    .setFontColor('#ffffff');


  indexSheet
    .setColumnWidth(1, 120);

  indexSheet
    .setColumnWidth(2, 80);

  indexSheet
    .setColumnWidth(3, 90);

  indexSheet
    .setColumnWidth(4, 350);

  indexSheet
    .setColumnWidth(5, 110);


  indexSheet
    .getDataRange()
    .setVerticalAlignment('top')
    .setWrap(true);


  return {

    success: true,
    count: result.length - 1,
    message: '位置索引建立完成'

  };

}



/**
 * 解析單一棟別區塊
 */
function parseLocationBlock(
  sheet,
  result,
  building,
  startRow,
  endRow,
  floorColumn,
  startPositionColumn,
  endPositionColumn,
  positionHeaderRow
) {

  for (
    let row = startRow;
    row <= endRow;
    row++
  ) {

    const floor =
      cleanText(
        sheet
          .getRange(
            row,
            floorColumn
          )
          .getDisplayValue()
      );


    if (!floor) {

      continue;

    }


    for (
      let column = startPositionColumn;
      column <= endPositionColumn;
      column++
    ) {

      const value =
        cleanText(
          sheet
            .getRange(
              row,
              column
            )
            .getDisplayValue()
        );


      if (!value) {

        continue;

      }


      const position =
        cleanText(
          sheet
            .getRange(
              positionHeaderRow,
              column
            )
            .getDisplayValue()
        );


      const cellAddress =
        sheet
          .getRange(
            row,
            column
          )
          .getA1Notation();


      result.push([

        building,
        floor,
        position,
        value,
        cellAddress

      ]);

    }

  }

}



/**
 * ==========================================================
 * PWA 呼叫此函式：
 * 廠商整併 + 位置配對
 * ==========================================================
 */
function getMergedVendorDataWithLocations() {

  const base =
    getMergedVendorData();


  if (!base || !base.success) {

    return base;

  }


  const vendors =
    base.data
    ||
    [];


  const locationResult =
    matchLocationsToVendors(
      vendors
    );


  base.data =
    locationResult.vendors;


  base.locationCount =
    locationResult.locationCount;


  base.matchedLocationCount =
    locationResult.matchedCount;


  base.unmatchedLocationCount =
    locationResult.unmatchedCount;


  base.ambiguousLocationCount =
    locationResult.ambiguousCount;


  return base;

}



/**
 * ==========================================================
 * 位置資料與正式廠商配對
 * ==========================================================
 */
function matchLocationsToVendors(vendors) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const locationSheet =
    ss.getSheetByName(
      '位置索引'
    );


  if (!locationSheet) {

    throw new Error(
      '找不到工作表：位置索引'
    );

  }


  // 每家廠商先建立空白位置陣列
  vendors.forEach(function(vendor) {

    vendor.locations = [];

  });


  const locations =
    readLocationIndex(
      locationSheet
    );


  // 第27步：建立「進駐位置圖」工作表與原始儲存格的精準連結
  const locationMapSheet =
    ss.getSheetByName(
      '進駐位置圖'
    );


  if (!locationMapSheet) {

    throw new Error(
      '找不到工作表：進駐位置圖'
    );

  }


  const spreadsheetId =
    ss.getId();


  const locationMapSheetId =
    locationMapSheet.getSheetId();


  locations.forEach(function(location) {

    location.sheetUrl =
      makeLocationSheetUrl(
        spreadsheetId,
        locationMapSheetId,
        location.cell
      );

  });


  const aliasMap =
    readLocationAliasMap();


  const vendorIndex =
    buildVendorSearchIndex(
      vendors
    );


  let matchedCount = 0;

  let unmatchedCount = 0;

  let ambiguousCount = 0;


  const report = [[

    '棟別',
    '樓層',
    '位置編號',
    '位置圖內容',
    '拆解名稱',
    '配對狀態',
    '正式廠商名稱',
    '統一編號',
    '配對方式',
    '原始儲存格'

  ]];


  locations.forEach(function(location) {

    const tokens =
      splitLocationNames(
        location.rawText
      );


    tokens.forEach(function(token) {

      const match =
        findVendorForLocationToken(
          token,
          vendors,
          vendorIndex,
          aliasMap
        );


      if (match.status === 'matched') {

        matchedCount++;


        const vendor =
          match.vendor;


        addVendorLocation(
          vendor,
          location,
          token,
          match.method
        );


        report.push([

          location.building,
          location.floor,
          location.position,
          location.rawText,
          token,
          '已配對',
          vendor.name,
          vendor.taxId,
          match.method,
          location.cell

        ]);

      }
      else if (
        match.status
        ===
        'ambiguous'
      ) {

        ambiguousCount++;


        report.push([

          location.building,
          location.floor,
          location.position,
          location.rawText,
          token,
          '多筆候選',
          match.candidates
            .map(function(v) {
              return v.name;
            })
            .join('、'),
          '',
          '需人工確認',
          location.cell

        ]);

      }
      else {

        unmatchedCount++;


        report.push([

          location.building,
          location.floor,
          location.position,
          location.rawText,
          token,
          '未配對',
          '',
          '',
          '請加入位置別名',
          location.cell

        ]);

      }

    });

  });



  // 建立每家廠商的主要位置資訊
  vendors.forEach(function(vendor) {

    if (
      vendor.locations.length
      >
      0
    ) {

      // 位置先依棟別、樓層、位置編號排序
      vendor.locations.sort(
        compareLocations
      );


      const first =
        vendor.locations[0];


      vendor.floor =
        first.floor
        ||
        '';


      vendor.position =
        first.position
        ||
        '';


      vendor.locationBuilding =
        first.building
        ||
        '';


      vendor.locationSummary =
        makeLocationSummary(
          first
        );

    }
    else {

      vendor.floor = '';

      vendor.position = '';

      vendor.locationBuilding = '';

      vendor.locationSummary = '';

    }

  });


  writeLocationMatchReport(
    report
  );


  return {

    vendors: vendors,
    locationCount: locations.length,
    matchedCount: matchedCount,
    unmatchedCount: unmatchedCount,
    ambiguousCount: ambiguousCount

  };

}



/**
 * 位置排序
 */
function compareLocations(a, b) {

  const buildingCompare =
    String(a.building || '')
      .localeCompare(
        String(b.building || ''),
        'zh-Hant'
      );


  if (buildingCompare !== 0) {

    return buildingCompare;

  }


  const floorA =
    locationFloorRank(
      a.floor
    );

  const floorB =
    locationFloorRank(
      b.floor
    );


  if (floorA !== floorB) {

    return floorB - floorA;

  }


  return String(
    a.position
    ||
    ''
  ).localeCompare(
    String(
      b.position
      ||
      ''
    ),
    'zh-Hant',
    {
      numeric: true
    }
  );

}



/**
 * Apps Script 端樓層排序值
 */
function locationFloorRank(floor) {

  const text =
    cleanText(floor)
      .toUpperCase();


  const basement =
    text.match(
      /^B(\d+)F?$/
    );


  if (basement) {

    return -
      Number(
        basement[1]
      );

  }


  const normal =
    text.match(
      /(\d+)/
    );


  if (normal) {

    return Number(
      normal[1]
    );

  }


  return -999;

}



/**
 * ==========================================================
 * 讀取「位置索引」
 * ==========================================================
 */
function readLocationIndex(sheet) {

  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (values.length < 2) {

    return [];

  }


  const headers =
    values[0].map(function(value) {

      return cleanText(value);

    });


  const column = {};


  headers.forEach(function(header, index) {

    column[header] = index;

  });


  const requiredHeaders = [

    '棟別',
    '樓層',
    '位置編號',
    '位置圖內容'

  ];


  requiredHeaders.forEach(function(header) {

    if (
      column[header]
      ===
      undefined
    ) {

      throw new Error(
        '位置索引缺少欄位：'
        +
        header
      );

    }

  });


  const result = [];


  for (
    let r = 1;
    r < values.length;
    r++
  ) {

    const row =
      values[r];


    const rawText =
      cleanText(
        row[
          column['位置圖內容']
        ]
      );


    if (!rawText) {

      continue;

    }


    result.push({

      building:
        cleanText(
          row[
            column['棟別']
          ]
        ),

      floor:
        cleanText(
          row[
            column['樓層']
          ]
        ),

      position:
        cleanText(
          row[
            column['位置編號']
          ]
        ),

      rawText:
        rawText,

      cell:
        column['原始儲存格']
        !==
        undefined
        ?
        cleanText(
          row[
            column['原始儲存格']
          ]
        )
        :
        ''

    });

  }


  return result;

}



/**
 * ==========================================================
 * 將位置圖內容拆成可能的廠商簡稱
 * ==========================================================
 */
function splitLocationNames(rawText) {

  let text =
    cleanText(rawText);


  if (!text) {

    return [];

  }


  text =
    text
      .replace(/\r/g, '\n')
      .replace(/；/g, '、')
      .replace(/;/g, '、')
      .replace(/，/g, '、')
      .replace(/,/g, '、');


  let parts = [];


  text
    .split(/\n+/)
    .forEach(function(line) {

      line =
        cleanText(line);


      if (!line) {

        return;

      }


      // 權億部分租：兆葳
      // 前後兩者都保留
      if (
        line.indexOf(
          '部分租'
        )
        >=
        0
      ) {

        const temp =
          line.split(
            /部分租[:：]?/
          );


        temp.forEach(function(value) {

          value =
            cleanLocationToken(
              value
            );


          if (value) {

            parts.push(value);

          }

        });


        return;

      }


      // 國城租:陶思 / 再望租:普明能
      // 以冒號後實際進駐者為主
      if (
        /租[:：]/.test(line)
      ) {

        const temp =
          line.split(
            /租[:：]/
          );


        const value =
          cleanLocationToken(
            temp[
              temp.length - 1
            ]
          );


        if (value) {

          parts.push(value);

        }


        return;

      }


      // 一般多家公司
      line
        .split('、')
        .forEach(function(value) {

          value =
            cleanLocationToken(
              value
            );


          if (value) {

            parts.push(value);

          }

        });

    });


  return Array.from(
    new Set(parts)
  );

}



/**
 * 清理位置圖簡稱
 */
function cleanLocationToken(value) {

  return cleanText(value)
    .replace(
      /^[\d\-\s]+/,
      ''
    )
    .replace(
      /^(國城|自用)[:：]?/,
      ''
    )
    .replace(
      /(進駐|承租)$/,
      ''
    )
    .trim();

}



/**
 * ==========================================================
 * 建立正式廠商名稱搜尋索引
 * ==========================================================
 */
function buildVendorSearchIndex(vendors) {

  const exact = {};


  vendors.forEach(function(vendor) {

    const names = [

      vendor.name,
      ...(vendor.otherNames || [])

    ];


    names.forEach(function(name) {

      const key =
        normalizeVendorNameForLocation(
          name
        );


      if (!key) {

        return;

      }


      if (!exact[key]) {

        exact[key] = [];

      }


      if (
        exact[key]
          .indexOf(vendor)
        ===
        -1
      ) {

        exact[key].push(vendor);

      }

    });

  });


  return {

    exact: exact

  };

}



/**
 * ==========================================================
 * 將位置圖簡稱配對正式廠商
 * ==========================================================
 */
function findVendorForLocationToken(
  token,
  vendors,
  vendorIndex,
  aliasMap
) {

  const tokenKey =
    normalizeVendorNameForLocation(
      token
    );


  if (!tokenKey) {

    return {
      status: 'unmatched'
    };

  }



  // 第一順位：人工位置別名
  if (aliasMap[tokenKey]) {

    const target =
      aliasMap[tokenKey];

    let candidate = null;


    if (target.taxId) {

      candidate =
        vendors.find(function(vendor) {

          return (
            vendor.taxIds
            ||
            []
          )
          .some(function(taxId) {

            return (
              normalizeTaxId(taxId)
              ===
              normalizeTaxId(
                target.taxId
              )
            );

          });

        });

    }


    if (
      !candidate
      &&
      target.name
    ) {

      const nameKey =
        normalizeVendorNameForLocation(
          target.name
        );


      candidate =
        vendors.find(function(vendor) {

          const names = [

            vendor.name,
            ...(vendor.otherNames || [])

          ];


          return names.some(function(name) {

            return (
              normalizeVendorNameForLocation(
                name
              )
              ===
              nameKey
            );

          });

        });

    }


    if (candidate) {

      return {
        status: 'matched',
        vendor: candidate,
        method: '位置別名'
      };

    }

  }



  // 第二順位：名稱完全相同
  const exactCandidates =
    vendorIndex
      .exact[tokenKey]
    ||
    [];


  if (
    exactCandidates.length
    ===
    1
  ) {

    return {
      status: 'matched',
      vendor: exactCandidates[0],
      method: '名稱完全比對'
    };

  }


  if (
    exactCandidates.length
    >
    1
  ) {

    return {
      status: 'ambiguous',
      candidates: exactCandidates
    };

  }



  // 第三順位：唯一簡稱包含比對
  if (
    tokenKey.length
    >=
    2
  ) {

    const candidates =
      vendors.filter(function(vendor) {

        const names = [

          vendor.name,
          ...(vendor.otherNames || [])

        ];


        return names.some(function(name) {

          const vendorKey =
            normalizeVendorNameForLocation(
              name
            );


          if (!vendorKey) {

            return false;

          }


          return (
            vendorKey.includes(tokenKey)
            ||
            tokenKey.includes(vendorKey)
          );

        });

      });


    const uniqueCandidates =
      Array.from(
        new Set(candidates)
      );


    if (
      uniqueCandidates.length
      ===
      1
    ) {

      return {
        status: 'matched',
        vendor: uniqueCandidates[0],
        method: '唯一簡稱比對'
      };

    }


    if (
      uniqueCandidates.length
      >
      1
    ) {

      return {
        status: 'ambiguous',
        candidates: uniqueCandidates
      };

    }

  }


  return {
    status: 'unmatched'
  };

}



/**
 * 公司名稱標準化（位置比對用）
 */
function normalizeVendorNameForLocation(value) {

  let text =
    cleanText(value);


  if (!text) {

    return '';

  }


  text =
    text
      .toLowerCase()
      .replace(/\s+/g, '')
      .replace(/　/g, '')
      .replace(/[()（）]/g, '')
      .replace(/股份有限公司/g, '')
      .replace(/有限責任公司/g, '')
      .replace(/有限公司/g, '')
      .replace(/財團法人/g, '')
      .replace(/社團法人/g, '')
      .replace(/高雄分公司/g, '')
      .replace(/高雄辦事處/g, '')
      .replace(/高雄辦公室/g, '')
      .replace(/高雄聯絡處/g, '')
      .replace(/辦事處/g, '')
      .replace(/辦公室/g, '')
      .replace(/聯絡處/g, '')
      .replace(/分公司/g, '');


  return text;

}



/**
 * 將位置加入廠商
 */
function addVendorLocation(
  vendor,
  location,
  matchedToken,
  matchMethod
) {

  const exists =
    vendor.locations.some(function(item) {

      return (
        item.building
        ===
        location.building
        &&
        item.floor
        ===
        location.floor
        &&
        item.position
        ===
        location.position
      );

    });


  if (exists) {

    return;

  }


  vendor.locations.push({

    building:
      location.building,

    floor:
      location.floor,

    position:
      location.position,

    rawText:
      location.rawText,

    matchedToken:
      matchedToken,

    matchMethod:
      matchMethod,

    sourceCell:
      location.cell,

    // 第27步：可直接開啟「進駐位置圖」並選取原始儲存格
    sheetUrl:
      location.sheetUrl
      ||
      ''

  });

}



/**
 * ==========================================================
 * 第27步：建立「進駐位置圖」精準儲存格網址
 *
 * 產生格式：
 * https://docs.google.com/spreadsheets/d/試算表ID/edit#gid=工作表ID&range=F24
 * ==========================================================
 */
function makeLocationSheetUrl(
  spreadsheetId,
  sheetId,
  cellAddress
) {

  const cell =
    cleanText(
      cellAddress
    );


  if (
    !spreadsheetId
    ||
    sheetId === ''
    ||
    sheetId === null
    ||
    sheetId === undefined
    ||
    !cell
  ) {

    return '';

  }


  return (
    'https://docs.google.com/spreadsheets/d/'
    +
    spreadsheetId
    +
    '/edit#gid='
    +
    sheetId
    +
    '&range='
    +
    encodeURIComponent(cell)
  );

}



/**
 * 位置摘要
 */
function makeLocationSummary(location) {

  if (!location) {

    return '';

  }


  const parts = [];


  if (location.building) {

    parts.push(
      location.building
    );

  }


  if (location.floor) {

    parts.push(
      location.floor
    );

  }


  if (location.position) {

    parts.push(
      '位置'
      +
      location.position
    );

  }


  return parts.join('｜');

}



/**
 * ==========================================================
 * 讀取「位置別名」
 * ==========================================================
 */
function readLocationAliasMap() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      '位置別名'
    );


  if (!sheet) {

    return {};

  }


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (values.length < 2) {

    return {};

  }


  const result = {};


  for (
    let r = 1;
    r < values.length;
    r++
  ) {

    const alias =
      cleanText(
        values[r][0]
      );

    const name =
      cleanText(
        values[r][1]
      );

    const taxId =
      cleanText(
        values[r][2]
      );


    if (!alias) {

      continue;

    }


    const key =
      normalizeVendorNameForLocation(
        alias
      );


    result[key] = {
      name: name,
      taxId: taxId
    };

  }


  return result;

}



/**
 * ==========================================================
 * 第一次建立「位置別名」工作表
 * ==========================================================
 */
function setupLocationAliasSheet() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  let sheet =
    ss.getSheetByName(
      '位置別名'
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        '位置別名'
      );

  }


  if (
    sheet.getLastRow()
    ===
    0
  ) {

    sheet
      .getRange(
        1,
        1,
        1,
        4
      )
      .setValues([[

        '位置圖名稱',
        '正式廠商名稱',
        '統一編號',
        '備註'

      ]]);


    sheet
      .getRange('A1:D1')
      .setBackground(
        '#075d97'
      )
      .setFontColor(
        '#ffffff'
      )
      .setFontWeight(
        'bold'
      );


    sheet
      .setFrozenRows(1);


    sheet
      .setColumnWidth(
        1,
        160
      );

    sheet
      .setColumnWidth(
        2,
        320
      );

    sheet
      .setColumnWidth(
        3,
        120
      );

    sheet
      .setColumnWidth(
        4,
        250
      );

  }


  return {
    success: true,
    message: '位置別名工作表已建立'
  };

}



/**
 * ==========================================================
 * 寫入「位置配對結果」
 * ==========================================================
 */
function writeLocationMatchReport(report) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  let sheet =
    ss.getSheetByName(
      '位置配對結果'
    );


  if (!sheet) {

    sheet =
      ss.insertSheet(
        '位置配對結果'
      );

  }


  sheet.clear();


  if (report.length === 0) {

    return;

  }


  sheet
    .getRange(
      1,
      1,
      report.length,
      report[0].length
    )
    .setValues(
      report
    );


  sheet
    .getRange(
      1,
      1,
      1,
      report[0].length
    )
    .setBackground(
      '#075d97'
    )
    .setFontColor(
      '#ffffff'
    )
    .setFontWeight(
      'bold'
    );


  sheet
    .setFrozenRows(1);


  sheet
    .getDataRange()
    .setVerticalAlignment(
      'top'
    )
    .setWrap(true);


  sheet
    .setColumnWidth(
      1,
      120
    );

  sheet
    .setColumnWidth(
      2,
      80
    );

  sheet
    .setColumnWidth(
      3,
      90
    );

  sheet
    .setColumnWidth(
      4,
      280
    );

  sheet
    .setColumnWidth(
      5,
      160
    );

  sheet
    .setColumnWidth(
      6,
      100
    );

  sheet
    .setColumnWidth(
      7,
      320
    );

}



/**
 * 手動重新產生位置配對報告
 */
function refreshLocationMatchReport() {

  const result =
    getMergedVendorDataWithLocations();


  return {

    success:
      true,

    mergedCount:
      result.mergedCount,

    matched:
      result.matchedLocationCount,

    unmatched:
      result.unmatchedLocationCount,

    ambiguous:
      result.ambiguousLocationCount

  };

}
