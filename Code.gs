// ============================================================
// Library Buddy — Google Apps Script Backend
// Five Elements International School — Team The Page Pioneer
// ============================================================

const SPREADSHEET_ID = '1KLm0CdoQLiY8Qmf4uI6eALIBCXtnGUNesFvYrNw4lIw';
const BOOKS_SHEET = 'Books';
const BORROWS_SHEET = 'Borrows';
const USERS_SHEET = 'Users';
const CATEGORIES_SHEET = 'Categories';

function getSheet(name) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toLowerCase() === name.trim().toLowerCase()) return sheets[i];
  }
  return null;
}

function sheetToArray(name) {
  var sheet = getSheet(name);
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return [];
  var h = data[0], rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < h.length; j++) obj[h[j]] = data[i][j];
    rows.push(obj);
  }
  return rows;
}

// --- Web App ---

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'books';
  var result;
  switch (action) {
    case 'books': result = sheetToArray(BOOKS_SHEET); break;
    case 'borrows': result = sheetToArray(BORROWS_SHEET); break;
    case 'categories': result = getList(CATEGORIES_SHEET); break;
    case 'overdue': result = getOverdue(); break;
    case 'stats': result = getStats(); break;
    case 'users': result = getUsers(e.parameter.u, e.parameter.p); break;
    default: result = {error:'Unknown'};
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'login') {
      return ContentService.createTextOutput(JSON.stringify(login(data.username, data.password))).setMimeType(ContentService.MimeType.JSON);
    }
    var auth = login(data.auth ? data.auth.username : '', data.auth ? data.auth.password : '');
    if (!auth.success || auth.role !== 'admin') {
      return ContentService.createTextOutput(JSON.stringify({success:false,message:'Unauthorized'})).setMimeType(ContentService.MimeType.JSON);
    }
    var result;
    switch (data.action) {
      case 'addBook': result = addBook(data); break;
      case 'updateBook': result = updateBook(data); break;
      case 'deleteBook': result = deleteBook(data); break;
      case 'borrow': result = borrowBook(data); break;
      case 'returnBook': result = returnBook(data); break;
      case 'saveCategories': result = saveList(CATEGORIES_SHEET, 'Category', data.items); break;
      case 'setupData': setupData(); result = {success:true}; break;
      default: result = {error:'Unknown action'};
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success:false,error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// --- Books CRUD ---

function addBook(data) {
  var sheet = getSheet(BOOKS_SHEET);
  var id = 'BK-' + Date.now().toString(36).toUpperCase();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
  sheet.appendRow([id, data.title||'', data.author||'', data.category||'', data.isbn||'', parseInt(data.totalCopies)||1, parseInt(data.totalCopies)||1, now, data.shelf||'']);
  return {success:true, id:id};
}

function updateBook(data) {
  var sheet = getSheet(BOOKS_SHEET);
  var all = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < all.length; i++) {
    if (all[i][0] === data.id) {
      var r = i + 1;
      if (data.title !== undefined) sheet.getRange(r,2).setValue(data.title);
      if (data.author !== undefined) sheet.getRange(r,3).setValue(data.author);
      if (data.category !== undefined) sheet.getRange(r,4).setValue(data.category);
      if (data.isbn !== undefined) sheet.getRange(r,5).setValue(data.isbn);
      if (data.totalCopies !== undefined) {
        var oldTotal = parseInt(all[i][5])||0;
        var oldAvail = parseInt(all[i][6])||0;
        var newTotal = parseInt(data.totalCopies)||0;
        var diff = newTotal - oldTotal;
        sheet.getRange(r,6).setValue(newTotal);
        sheet.getRange(r,7).setValue(Math.max(0, oldAvail + diff));
      }
      if (data.shelf !== undefined) sheet.getRange(r,9).setValue(data.shelf);
      return {success:true};
    }
  }
  return {success:false, message:'Not found'};
}

function deleteBook(data) {
  var sheet = getSheet(BOOKS_SHEET);
  var all = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < all.length; i++) {
    if (all[i][0] === data.id) { sheet.deleteRow(i+1); return {success:true}; }
  }
  return {success:false, message:'Not found'};
}

// --- Borrow / Return ---

function borrowBook(data) {
  var booksSheet = getSheet(BOOKS_SHEET);
  var books = booksSheet.getDataRange().getDisplayValues();
  for (var i = 1; i < books.length; i++) {
    if (books[i][0] === data.bookId) {
      var avail = parseInt(books[i][6]) || 0;
      if (avail <= 0) return {success:false, message:'No copies available'};
      booksSheet.getRange(i+1, 7).setValue(avail - 1);

      var borrowSheet = getSheet(BORROWS_SHEET);
      if (!borrowSheet) {
        borrowSheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(BORROWS_SHEET);
        borrowSheet.appendRow(['BorrowID','BookID','BookTitle','Borrower','Grade','BorrowDate','DueDate','ReturnDate','Status']);
      }
      var bid = 'BR-' + Date.now().toString(36).toUpperCase();
      var now = new Date();
      var due = new Date(now.getTime() + (parseInt(data.days)||14) * 86400000);
      var borrowDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd MMM yyyy');
      var dueDate = Utilities.formatDate(due, Session.getScriptTimeZone(), 'dd MMM yyyy');
      borrowSheet.appendRow([bid, data.bookId, books[i][1], data.borrower||'', data.grade||'', borrowDate, dueDate, '', 'Borrowed']);
      return {success:true, borrowId:bid, dueDate:dueDate};
    }
  }
  return {success:false, message:'Book not found'};
}

function returnBook(data) {
  var borrowSheet = getSheet(BORROWS_SHEET);
  if (!borrowSheet) return {success:false, message:'No borrows sheet'};
  var borrows = borrowSheet.getDataRange().getDisplayValues();
  var h = borrows[0];
  var idCol = 0, bookIdCol = h.indexOf('BookID'), statusCol = h.indexOf('Status'), returnCol = h.indexOf('ReturnDate');

  for (var i = 1; i < borrows.length; i++) {
    if (borrows[i][idCol] === data.borrowId && borrows[i][statusCol] === 'Borrowed') {
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
      borrowSheet.getRange(i+1, returnCol+1).setValue(now);
      borrowSheet.getRange(i+1, statusCol+1).setValue('Returned');

      // Increment available copies
      var booksSheet = getSheet(BOOKS_SHEET);
      var books = booksSheet.getDataRange().getDisplayValues();
      for (var j = 1; j < books.length; j++) {
        if (books[j][0] === borrows[i][bookIdCol]) {
          var avail = parseInt(books[j][6]) || 0;
          booksSheet.getRange(j+1, 7).setValue(avail + 1);
          break;
        }
      }
      return {success:true};
    }
  }
  return {success:false, message:'Borrow record not found'};
}

// --- Overdue ---

function getOverdue() {
  var borrows = sheetToArray(BORROWS_SHEET);
  var now = new Date();
  return borrows.filter(function(b) {
    if (b.Status !== 'Borrowed') return false;
    var due = new Date(b.DueDate);
    return now > due;
  });
}

// --- Stats ---

function getStats() {
  var books = sheetToArray(BOOKS_SHEET);
  var borrows = sheetToArray(BORROWS_SHEET);
  var totalBooks = books.length;
  var totalCopies = 0, totalAvail = 0;
  var catCount = {};
  for (var i = 0; i < books.length; i++) {
    totalCopies += parseInt(books[i].TotalCopies)||0;
    totalAvail += parseInt(books[i].Available)||0;
    var c = books[i].Category||'Other';
    catCount[c] = (catCount[c]||0) + 1;
  }
  var activeBorrows = borrows.filter(function(b){return b.Status==='Borrowed';}).length;
  var overdue = getOverdue().length;
  return {totalBooks:totalBooks, totalCopies:totalCopies, available:totalAvail, activeBorrows:activeBorrows, overdue:overdue, categories:catCount};
}

// --- Helpers ---

function getList(sheetName) {
  var sheet = getSheet(sheetName);
  if (!sheet) return [];
  var data = sheet.getDataRange().getDisplayValues();
  var r = [];
  for (var i = 1; i < data.length; i++) if (data[i][0]) r.push(data[i][0].trim());
  return r;
}

function saveList(sheetName, header, items) {
  var sheet = getSheet(sheetName) || SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(sheetName);
  sheet.clear(); sheet.appendRow([header]);
  (items||[]).forEach(function(v){sheet.appendRow([v]);});
  return {success:true};
}

function login(username, password) {
  var sheet = getSheet(USERS_SHEET);
  if (!sheet) return {success:false, message:'Users sheet not found'};
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim()===String(username).trim() && String(data[i][1]).trim()===String(password).trim())
      return {success:true, role:String(data[i][2]).trim(), displayName:String(data[i][3]).trim(), username:String(data[i][0]).trim()};
  }
  return {success:false, message:'Invalid credentials'};
}

function getUsers(u, p) {
  var auth = login(u, p);
  if (!auth.success || auth.role !== 'admin') return {error:'Unauthorized'};
  var sheet = getSheet(USERS_SHEET);
  var data = sheet.getDataRange().getValues(), r = [];
  for (var i = 1; i < data.length; i++) if (data[i][0]) r.push({username:String(data[i][0]),password:String(data[i][1]),role:String(data[i][2]),displayName:String(data[i][3])});
  return r;
}

// --- Setup ---

function setupData() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var us = ss.getSheetByName(USERS_SHEET)||ss.insertSheet(USERS_SHEET); us.clear();
  us.appendRow(['Username','Password','Role','DisplayName']);
  us.appendRow(['admin','admin123','admin','Administrator']);
  us.appendRow(['prateeksha','teach123','admin','Ms. Prateeksha']);

  var cs = ss.getSheetByName(CATEGORIES_SHEET)||ss.insertSheet(CATEGORIES_SHEET); cs.clear();
  cs.appendRow(['Category']);
  ['Fiction','Non-Fiction','Science','Mathematics','History','Geography','Literature','Comics','Reference','Biography','Poetry','Technology'].forEach(function(c){cs.appendRow([c]);});

  var bs = ss.getSheetByName(BOOKS_SHEET)||ss.insertSheet(BOOKS_SHEET); bs.clear();
  bs.appendRow(['ID','Title','Author','Category','ISBN','TotalCopies','Available','DateAdded','Shelf']);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMM yyyy');
  var samples = [
    ['The Jungle Book','Rudyard Kipling','Fiction','','3','3','A1'],
    ['A Brief History of Time','Stephen Hawking','Science','','2','2','B2'],
    ['Wings of Fire','APJ Abdul Kalam','Biography','','4','4','C1'],
    ['Harry Potter and the Philosopher\'s Stone','J.K. Rowling','Fiction','','3','3','A2'],
    ['Mathematics for Class 9','R.D. Sharma','Mathematics','','10','10','D1'],
    ['Discovery of India','Jawaharlal Nehru','History','','2','2','C3'],
    ['Tinkle Comics Collection','Anant Pai','Comics','','5','5','E1'],
    ['Oxford English Dictionary','Oxford Press','Reference','','3','3','F1'],
    ['The Story of My Experiments with Truth','Mahatma Gandhi','Biography','','2','2','C2'],
    ['Panchatantra Stories','Vishnu Sharma','Fiction','','4','4','A3'],
  ];
  for (var i = 0; i < samples.length; i++) {
    var s = samples[i];
    var id = 'BK-' + (Date.now()+i).toString(36).toUpperCase();
    bs.appendRow([id, s[0], s[1], s[2], s[3], s[4], s[5], now, s[6]]);
  }

  var br = ss.getSheetByName(BORROWS_SHEET)||ss.insertSheet(BORROWS_SHEET); br.clear();
  br.appendRow(['BorrowID','BookID','BookTitle','Borrower','Grade','BorrowDate','DueDate','ReturnDate','Status']);

  Logger.log('Setup complete');
}
