const XLSX = require('xlsx');
const fs = require('fs');
const html = fs.readFileSync('/Users/presenter/Desktop/An/dashboard.html', 'utf8');

// Extract parseData function logic manually to test it
const data = [
  {
     "Store ID": "ID001",
     "Store Name": "Mega Bangna",
     "Staff ID": "S123",
     "Staff Name": "John",
     "Completed Date (Day)": "4/12/2026",
     "Transaction Type": "AAR Mac 1:1 Demo",
     "ลูกค้าได้รับการนำเสนอข้อมูลสินค้า Apple ในเรื่องใด?": "" // Blank
  },
  {
     "Store ID": "ID001",
     "Store Name": "Mega Bangna",
     "Staff ID": "S124",
     "Staff Name": "Jane",
     "Completed Date (Day)": "4/12/2026",
     "Transaction Type": "Normal Sale",
     "ลูกค้าได้รับการนำเสนอข้อมูลสินค้า Apple ในเรื่องใด?": "Mac" // Populated
  }
];

let state = { rating: [], sales: [], fileNames: {}, availableYears: new Set(), availableMonths: new Set(), availableBranches: new Set(), availableCats: new Set() };
let isSales = false;
let isRating = true;
let isRating30 = true;

const headers = Object.keys(data[0]);
const getCol = (kws) => { let c = headers.find(h => kws.some(k => h.toLowerCase() === k.toLowerCase())); if (c) return c; return headers.find(h => kws.some(k => h.toLowerCase().includes(k.toLowerCase()))); };

function classifyCategory(cat, subCat, mwad, catId) {
  cat = (cat||'').toLowerCase();
  if (cat.includes('mac')) return 'Mac';
  if (cat.includes('ipad')) return 'iPad';
  if (cat.includes('iphone')) return 'iPhone';
  if (cat.includes('watch')) return 'Apple Watch';
  return 'Accessories & Others';
}

function getVal(val) { return isNaN(parseFloat(val)) ? 0 : parseFloat(val); }
function normalizeDate(d) { return new Date(); } // Mock

const staffNameKey = getCol(['officer (name)', 'staff name', 'พนักงาน']);
const staffIdKey = getCol(['officer (id)', 'staff id', 'รหัส']);
const transactionTypeKey = getCol(['transaction type', 'transaction_type', 'transaction', 'ประเภทรายการ']) || headers[14];
const topicKey = getCol(['ลูกค้าได้รับการนำเสนอข้อมูลสินค้า Apple ในเรื่องใด?', 'เรื่องใด', 'topic']);

data.forEach(row => {
  let mainTopic = String(row[topicKey] || '').trim();
  let txType = String(row[transactionTypeKey] || '').trim();
  let combinedStr = (mainTopic + " " + txType).toLowerCase();
  
  let rawCat = '';
  if (combinedStr.includes('mac')) rawCat = 'Mac';
  else if (combinedStr.includes('ipad')) rawCat = 'iPad';
  else if (combinedStr.includes('iphone')) rawCat = 'iPhone';
  else if (combinedStr.includes('watch')) rawCat = 'Apple Watch';
  else {
      let match = txType.match(/AAR\s+(.+?)\s*(?:1:1|Demo)/i);
      rawCat = (match && match[1]) ? match[1].trim() : (mainTopic || txType);
  }
  
  let topicStr = mainTopic || rawCat; 
  let groupCat = classifyCategory(rawCat, '', '', '');
  state.availableCats.add(groupCat);

  const obj = {
    staffName: String(row[staffNameKey] || 'ไม่ระบุ').trim(),
    staffId: String(row[staffIdKey] || '-').trim(),
    category: groupCat,
    topic: isRating ? (String(row[topicKey] || '').trim() || topicStr || 'อื่นๆ') : 'อื่นๆ'
  };

  if (obj.staffName) {
      if (isRating) {
          if (!isSales || obj.rating > 0 || (obj.topic !== 'อื่นๆ' && obj.topic !== '')) {
              state.rating.push(obj);
          }
      }
  }
});

console.log("State Rating Array:", JSON.stringify(state.rating, null, 2));
console.log("Filter match for Mac:", state.rating.filter(x => x.category === 'Mac').length);

