/**
 * ==========================================================
 * 高雄軟體園區廠商查詢系統
 * Code.gs
 * Version 7.0 - 第27步
 *
 * 功能：
 * 1. 顯示手機查詢首頁
 * 2. 讀取「區內事業」
 * 3. 讀取「營業或聯絡處所」
 * 4. 自動整併重複廠商
 * 5. 提供 JSON API
 * 6. 供 Location.gs 加入進駐位置資料
 * 7. 第27步支援 PWA 精準連結「進駐位置圖」原始儲存格
 * ==========================================================
 */


/**
 * Web App 入口
 */
function doGet(e) {

  // 沒有 action -> 顯示查詢首頁
  if (!e || !e.parameter || !e.parameter.action) {

    return HtmlService
      .createTemplateFromFile('Index')
      .evaluate()
      .setTitle('高雄軟體園區廠商查詢')
      .addMetaTag(
        'viewport',
        'width=device-width, initial-scale=1, viewport-fit=cover'
      );

  }


  try {

    const action = e.parameter.action;


    switch (action) {

      case 'status':

        return outputJSON({
          success: true,
          api: '高雄軟體園區廠商查詢API',
          version: '7.0',
          message: 'API 運作正常'
        });


      case 'companies': {

        const data = getSheetData('區內事業');

        return outputJSON({
          success: true,
          type: 'companies',
          count: data.length,
          data: data
        });

      }


      case 'offices': {

        const data = getSheetData('營業或聯絡處所');

        return outputJSON({
          success: true,
          type: 'offices',
          count: data.length,
          data: data
        });

      }


      case 'merged':

        return outputJSON(
          getMergedVendorData()
        );


      case 'mergedLocations':

        return outputJSON(
          getMergedVendorDataWithLocations()
        );


      default:

        return outputJSON({
          success: false,
          message: '未知的 action：' + action
        });

    }

  }
  catch (error) {

    return outputJSON({
      success: false,
      message: error.message,
      stack: error.stack || ''
    });

  }

}



/**
 * ==========================================================
 * 整併兩張廠商資料表
 * ==========================================================
 */
function getMergedVendorData() {

  const companies =
    getSheetData('區內事業');

  const offices =
    getSheetData('營業或聯絡處所');


  const vendorList = [];

  const taxIndex = {};

  const nameIndex = {};


  /**
   * 尋找或建立廠商
   */
  function findOrCreateVendor(name, taxId) {

    const taxKey =
      normalizeTaxId(taxId);

    const nameKey =
      normalizeCompanyName(name);

    let vendor = null;


    // 第一順位：統一編號
    if (taxKey && taxIndex[taxKey]) {

      vendor =
        taxIndex[taxKey];

    }


    // 第二順位：完整公司名稱
    if (!vendor && nameKey && nameIndex[nameKey]) {

      vendor =
        nameIndex[nameKey];

    }


    // 找不到就建立
    if (!vendor) {

      vendor = {

        names: [],
        taxIds: [],
        sources: [],

        companyAddresses: [],
        companyBuildings: [],
        companyPhones: [],
        businessItems: [],

        officeAddresses: [],
        officeBuildings: [],
        officePhones: [],
        industryTypes: [],
        postalCodes: []

      };

      vendorList.push(vendor);

    }


    // 建立索引
    if (taxKey) {

      taxIndex[taxKey] = vendor;

    }

    if (nameKey) {

      nameIndex[nameKey] = vendor;

    }


    return vendor;

  }



  // ========================================================
  // 區內事業
  // ========================================================
  companies.forEach(function(item) {

    const name =
      cleanText(
        item['廠商名稱']
      );

    const taxId =
      cleanText(
        item['統一編號/工廠登記證號']
      );


    if (!name) {

      return;

    }


    const vendor =
      findOrCreateVendor(
        name,
        taxId
      );


    addUnique(
      vendor.names,
      name
    );

    addUnique(
      vendor.taxIds,
      taxId
    );

    addUnique(
      vendor.sources,
      '區內事業'
    );

    addUnique(
      vendor.companyAddresses,
      item['地址']
    );

    addUnique(
      vendor.companyBuildings,
      item['棟別']
    );

    addUnique(
      vendor.companyPhones,
      item['電話']
    );

    addUnique(
      vendor.businessItems,
      item['營業項目']
    );

  });



  // ========================================================
  // 營業或聯絡處所
  // ========================================================
  offices.forEach(function(item) {

    const name =
      cleanText(
        item['廠商名稱']
      );

    const taxId =
      cleanText(
        item['統一編號']
      );


    if (!name) {

      return;

    }


    const vendor =
      findOrCreateVendor(
        name,
        taxId
      );


    addUnique(
      vendor.names,
      name
    );

    addUnique(
      vendor.taxIds,
      taxId
    );

    addUnique(
      vendor.sources,
      '營業或聯絡處所'
    );

    addUnique(
      vendor.officeAddresses,
      item['營業處地址']
    );

    addUnique(
      vendor.officeBuildings,
      item['棟別']
    );

    addUnique(
      vendor.officePhones,
      item['電話']
    );

    addUnique(
      vendor.industryTypes,
      item['行業別']
    );

    addUnique(
      vendor.postalCodes,
      item['郵遞區號']
    );

  });



  // ========================================================
  // 整理成前端使用格式
  // ========================================================
  const result =
    vendorList.map(function(vendor) {

      const building =
        vendor.companyBuildings[0]
        ||
        vendor.officeBuildings[0]
        ||
        '';

      const address =
        vendor.companyAddresses[0]
        ||
        vendor.officeAddresses[0]
        ||
        '';

      const phone =
        vendor.companyPhones[0]
        ||
        vendor.officePhones[0]
        ||
        '';


      let sourceLabel = '';

      if (vendor.sources.length >= 2) {

        sourceLabel = '雙資料來源';

      }
      else {

        sourceLabel =
          vendor.sources[0]
          ||
          '';

      }


      return {

        name:
          vendor.names[0]
          ||
          '',

        otherNames:
          vendor.names.slice(1),

        taxId:
          vendor.taxIds[0]
          ||
          '',

        taxIds:
          vendor.taxIds,

        sources:
          vendor.sources,

        sourceLabel:
          sourceLabel,

        building:
          building,

        address:
          address,

        phone:
          phone,

        companyAddresses:
          vendor.companyAddresses,

        companyBuildings:
          vendor.companyBuildings,

        companyPhones:
          vendor.companyPhones,

        businessItems:
          vendor.businessItems,

        officeAddresses:
          vendor.officeAddresses,

        officeBuildings:
          vendor.officeBuildings,

        officePhones:
          vendor.officePhones,

        industryTypes:
          vendor.industryTypes,

        postalCodes:
          vendor.postalCodes

      };

    });


  // 公司名稱排序
  result.sort(function(a, b) {

    return a.name.localeCompare(
      b.name,
      'zh-Hant'
    );

  });


  return {

    success:
      true,

    rawCount:
      companies.length
      +
      offices.length,

    companyCount:
      companies.length,

    officeCount:
      offices.length,

    mergedCount:
      result.length,

    duplicateRemoved:
      companies.length
      +
      offices.length
      -
      result.length,

    data:
      result

  };

}



/**
 * ==========================================================
 * 讀取 Google 試算表
 * ==========================================================
 */
function getSheetData(sheetName) {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      sheetName
    );


  if (!sheet) {

    throw new Error(
      '找不到工作表：' + sheetName
    );

  }


  const values =
    sheet
      .getDataRange()
      .getDisplayValues();


  if (values.length < 2) {

    return [];

  }


  const headers =
    values[0].map(function(header) {

      return cleanText(header);

    });


  const data = [];


  for (
    let r = 1;
    r < values.length;
    r++
  ) {

    const row =
      values[r];


    const hasData =
      row.some(function(value) {

        return cleanText(value) !== '';

      });


    if (!hasData) {

      continue;

    }


    const item = {};


    headers.forEach(function(header, c) {

      if (header) {

        item[header] =
          row[c]
          ||
          '';

      }

    });


    data.push(item);

  }


  return data;

}



/**
 * 清理文字
 */
function cleanText(value) {

  return String(
    value
    ||
    ''
  ).trim();

}



/**
 * 標準化統一編號
 */
function normalizeTaxId(value) {

  return cleanText(value)
    .replace(/\s+/g, '')
    .replace(/'/g, '')
    .toUpperCase();

}



/**
 * 標準化公司名稱
 */
function normalizeCompanyName(value) {

  return cleanText(value)
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/　/g, '')
    .replace(/[()（）]/g, '');

}



/**
 * 陣列只加入唯一且非空白值
 */
function addUnique(array, value) {

  const text =
    cleanText(value);


  if (!text) {

    return;

  }


  if (
    array.indexOf(text)
    ===
    -1
  ) {

    array.push(text);

  }

}



/**
 * JSON 輸出
 */
function outputJSON(data) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService
        .MimeType
        .JSON
    );

}
