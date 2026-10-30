/** STOCKLY — Google Apps Script. Crie uma planilha nova e cole o ID abaixo. */
const SPREADSHEET_ID = 'COLE_AQUI_O_ID_DA_PLANILHA';
const SHEETS = { products: 'Produtos', movements: 'Movimentacoes' };
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheet_(ss, SHEETS.products, ['id','name','category','unit','quantity','min','cost','updatedAt']);
  ensureSheet_(ss, SHEETS.movements, ['id','date','type','productId','productName','quantity','note','user']);
  return 'Planilha preparada. Agora publique o web app.';
}
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) { sh.appendRow(headers); sh.setFrozenRows(1); sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#173321').setFontColor('#ffffff'); }
}
function doGet() { return json_({ok:true,data:{message:'Stockly API online'}}); }
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const result = route_(ss, body);
    return json_({ok:true,data:result});
  } catch(err) { return json_({ok:false,error:String(err.message || err)}); }
  finally { try { lock.releaseLock(); } catch(ignore) {} }
}
function route_(ss, b) {
  if (b.action === 'getData') return { products: readObjects_(ss.getSheetByName(SHEETS.products)), movements: readObjects_(ss.getSheetByName(SHEETS.movements)).reverse() };
  if (b.action === 'saveProduct') return saveProduct_(ss, b.product || {});
  if (b.action === 'deleteProduct') return deleteProduct_(ss, b.id);
  if (b.action === 'moveStock') return moveStock_(ss, b);
  throw new Error('Ação não reconhecida.');
}
function readObjects_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues(), headers = values.shift();
  return values.filter(r => r[0] !== '').map(r => { const o={}; headers.forEach((h,i)=>o[h]=r[i] instanceof Date?r[i].toISOString():r[i]); return o; });
}
function saveProduct_(ss, p) {
  if (!p.name || !String(p.name).trim()) throw new Error('Informe o nome do produto.');
  const sh=ss.getSheetByName(SHEETS.products), rows=readObjects_(sh); const id=p.id || Utilities.getUuid();
  const row=[id,String(p.name).trim(),p.category||'',p.unit||'un.',Number(p.quantity)||0,Number(p.min)||0,Number(p.cost)||0,new Date()];
  const idx=rows.findIndex(x=>String(x.id)===String(id));
  if(idx>=0) sh.getRange(idx+2,1,1,row.length).setValues([row]); else sh.appendRow(row);
  return {id:id};
}
function deleteProduct_(ss,id) {
  const sh=ss.getSheetByName(SHEETS.products), rows=readObjects_(sh), idx=rows.findIndex(x=>String(x.id)===String(id));
  if(idx<0) throw new Error('Produto não encontrado.'); sh.deleteRow(idx+2); return {deleted:true};
}
function moveStock_(ss,b) {
  const type=b.type, amount=Number(b.quantity);
  if(!['entrada','saida'].includes(type)) throw new Error('Tipo de movimentação inválido.');
  if(!amount || amount<=0) throw new Error('Informe uma quantidade válida.');
  const sh=ss.getSheetByName(SHEETS.products), rows=readObjects_(sh), idx=rows.findIndex(x=>String(x.id)===String(b.productId));
  if(idx<0) throw new Error('Produto não encontrado.');
  const p=rows[idx], current=Number(p.quantity)||0;
  if(type==='saida' && amount>current) throw new Error('Estoque insuficiente. Disponível: '+current+' '+p.unit);
  const next=type==='entrada'?current+amount:current-amount;
  sh.getRange(idx+2,5).setValue(next); sh.getRange(idx+2,8).setValue(new Date());
  ss.getSheetByName(SHEETS.movements).appendRow([Utilities.getUuid(),new Date(),type,p.id,p.name,amount,b.note||'',Session.getActiveUser().getEmail()||'Usuário']);
  return {quantity:next};
}
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
