
// TP Sales Dashboard — Data API
// Deploy as: Web App → Execute as: Me → Who has access: Anyone

function doGet(e) {
  const SHEET_ID = '18UjuOiSNpJaTmL_Cad3Fc81Nt2oP4pOyTEdkvWtGtwg';
  const wb = SpreadsheetApp.openById(SHEET_ID);
  const SHEETS = ["Jan'26", "Feb'26", "Mar'26"];
  const KEYS   = ["Jan26",  "Feb26",  "Mar26"];
  
  const BDM_SLABS = ['Regular','Acting AM','Device Tgt','PIP','Warning','New','Institution'];
  const result = {};

  SHEETS.forEach(function(sheetName, si) {
    const sheet = wb.getSheetByName(sheetName);
    if (!sheet) return;
    const rows = sheet.getDataRange().getValues();
    
    var bdm=[], inside=[], am=[], rm=[], products=[], summary=null;
    var amHdr=-1, rmHdr=-1, prodHdr=-1;

    // Find section header rows
    for (var i=0; i<rows.length; i++) {
      var c2 = String(rows[i][2]||'').trim();
      if (c2==='AM' && !String(rows[i][3]||'').trim()) amHdr=i;
      if (c2==='RM' && !String(rows[i][3]||'').trim()) rmHdr=i;
      if (c2==='Product') prodHdr=i;
    }

    function pn(v){ var n=parseFloat(String(v||'').replace(/,/g,'')); return isNaN(n)?0:n; }

    for (var i=1; i<rows.length; i++) {
      var r=rows[i];
      var slab=String(r[0]||'').trim(), amTag=String(r[1]||'').trim();
      var srNo=String(r[2]||'').trim(), name=String(r[3]||'').trim(), loc=String(r[4]||'').trim();

      if (amHdr>0 && i<amHdr) {
        if (BDM_SLABS.indexOf(slab)>=0 && name) {
          bdm.push({name:name,am:amTag,location:loc,slab:slab,
            target:pn(r[5]),saas:pn(r[7]),cp:pn(r[8]),opd:pn(r[9]),
            dx:pn(r[10]),pillup:pn(r[11]),nutra:pn(r[12]),strip:pn(r[13]),cgm:pn(r[14]),
            coreRevenue:pn(r[15]),transactionRevenue:pn(r[16]),
            otherRevenue:pn(r[17]),totalRevenue:pn(r[18]),coreContributionPct:pn(r[19])});
        }
        if (slab==='Inside Sales' && name) {
          inside.push({name:name,am:amTag,location:loc,slab:slab,
            target:pn(r[5]),coreRevenue:pn(r[15]),transactionRevenue:pn(r[16]),
            otherRevenue:pn(r[17]),totalRevenue:pn(r[18]),coreContributionPct:pn(r[19])});
        }
        if (slab==='-' && name==='Total Revenue') {
          summary={target:pn(r[5]),coreRevenue:pn(r[15]),
            transactionRevenue:pn(r[16]),otherRevenue:pn(r[17]),totalRevenue:pn(r[18])};
        }
      }
      if (amHdr>0 && rmHdr>0 && i>amHdr && i<rmHdr && slab!=='# of BDMs' && name && name!=='Total Revenue') {
        am.push({name:name,rm:amTag,location:loc,numBDMs:parseInt(slab)||0,
          target:pn(r[5]),cgm:pn(r[14]),coreRevenue:pn(r[15]),transactionRevenue:pn(r[16]),
          otherRevenue:pn(r[17]),totalRevenue:pn(r[18]),revenuePerBDM:pn(r[19])});
      }
      if (rmHdr>0 && i>rmHdr && (prodHdr<0||i<prodHdr) && slab!=='# of BDMs' && name && name!=='Total Revenue') {
        rm.push({name:name,rm:amTag,location:loc,numBDMs:parseInt(slab)||0,
          target:pn(r[5]),cgm:pn(r[14]),coreRevenue:pn(r[15]),transactionRevenue:pn(r[16]),
          otherRevenue:pn(r[17]),totalRevenue:pn(r[18]),revenuePerBDM:pn(r[19])});
      }
      if (prodHdr>0 && i>prodHdr) {
        var pname=String(r[2]||'').trim();
        if (pname && pname!=='Product' && pname!=='Total') {
          var units=pn(r[3]), rev=pn(r[4]), aov=pn(r[5]);
          if (units>0||rev>0) products.push({product:pname,units:Math.round(units),revenue:rev,aov:aov});
        }
      }
    }

    if (!summary) {
      var all=[].concat(bdm,inside);
      summary={target:all.reduce(function(a,b){return a+b.target;},0),
        coreRevenue:all.reduce(function(a,b){return a+b.coreRevenue;},0),
        transactionRevenue:all.reduce(function(a,b){return a+b.transactionRevenue;},0),
        otherRevenue:all.reduce(function(a,b){return a+b.otherRevenue;},0),
        totalRevenue:all.reduce(function(a,b){return a+b.totalRevenue;},0)};
    }
    result[KEYS[si]] = {bdm:bdm,inside:inside,am:am,rm:rm,summary:summary,products:products};
  });

  var output = ContentService.createTextOutput(JSON.stringify(result));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
