const xlsx = require("xlsx");
const fs = require("fs");

// 讀取 Excel 檔案
const readExcelFile = (filePath) => {
  try {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    return xlsx.utils.sheet_to_json(sheet);
  } catch (error) {
    console.error("讀取或解析 Excel 檔案失敗:", error.message);
    process.exit(1);
  }
};

// 處理資料邏輯
const processItem = (item, index) => {
  console.log(`處理第 ${index + 1} 筆資料:`, item);

  // 初始化多張圖片
  const imagesUrl = item.imagesUrl
    ? item.imagesUrl.split(",").map((url) => url.trim())
    : [];

  return {
    title: item.title || "未提供標題",
    category: item.category || "未分類",
    description: item.description || "",
    content: item.content || "",
    origin_price: parseFloat(item.origin_price) || 0,
    price: parseFloat(item.price) || 0,
    unit: item.unit || "未定義",
    is_enabled: parseInt(item.is_enabled, 10) || 0,
    imageUrl: item.imageUrl || "",
    imagesUrl,
  };
};

// 將處理後的資料寫入 JSON 檔案
const writeJsonFile = (filePath, data) => {
  try {
    fs.writeFileSync(filePath, JSON.stringify(data, null, 2), "utf8");
    console.log(`資料轉換完成，已生成 ${filePath}`);
  } catch (error) {
    console.error("寫入 JSON 檔案失敗:", error.message);
  }
};

// 主程式執行
const main = () => {
  const inputFilePath = "products.xlsx";
  const outputFilePath = "products.json";

  const data = readExcelFile(inputFilePath);
  const processedData = data.map((item, index) => processItem(item, index));
  writeJsonFile(outputFilePath, processedData);
};

main();
