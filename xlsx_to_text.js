var xlsx2txt,cell2txt,a1ref;
cell2txt = function(cell, sst, styles) {
  var val, t, s, fmtId;
  val = cell.getElementsByTagName('v')[0];
  val = val ? val.textContent : null;
  if (val === null) {return null;}
  t = cell.getAttribute('t');
  s = cell.getAttribute('s')-0;
  if (t === 's') {return sst[val];}
  if (s && styles[s]) {
    fmtId = styles[s].numFmtId;
    if (fmtId >= 14 && fmtId < 37 || fmtId > 44 && fmtId < 48 || fmtId > 49 && fmtId < 59) {
      val = new Date((val - 25567) * 86400000);
    }
  }
  return val;
};

a1ref = function(a1) {
  var m, r, c, i, p;
  if (!(m = a1.toUpperCase().match(/^([A-Z]+)([0-9]+)$/))) {
    return null;
  }
  r = m[2]-1;
  for(c=0,i=m[1].length-1,p=1; i>=0; p*=26,i--) {
    c += (m[1].charCodeAt(i)-64) * p;
  }
  return [r, c-1];
};
xlsx2txt = function(file, callback) {
  var fr = new FileReader();
  fr.onload = function() {
    var zip,xml,dom,sst,styles,data,rows,cells,pos,txt,i,j;
    zip = new JSZip(fr.result);
    dom = new DOMParser();
    xml = {
      styles: dom.parseFromString(zip.file('xl/styles.xml').asText(), 'application/xml'),
      sst: dom.parseFromString(zip.file('xl/sharedStrings.xml').asText(), 'application/xml'),
      sheet1: dom.parseFromString(zip.file('xl/worksheets/sheet1.xml').asText(), 'application/xml')
    };
    txt = "";
    sst = {list:xml.sst.getElementsByTagName('si')};
    for(i=0; i<sst.list.length; i++) {
      sst[i] = sst.list[i].textContent;
    }
    styles = {list:xml.styles.getElementsByTagName('cellXfs')[0].childNodes};
    for(i=0; i<styles.list.length; i++) {
      styles[i] = {numFmtId: styles.list[i].getAttribute('numFmtId')-0};
    }
    data = [];
    rows = xml.sheet1.getElementsByTagName('row');
    for(i=0; i<rows.length; i++) {
      cells = rows[i].getElementsByTagName('c');
      for(j=0; j<cells.length; j++) {
        pos = a1ref(cells[j].getAttribute('r'));
        if (!data[pos[0]]) {data[pos[0]] = [];}
        data[pos[0]][pos[1]] = cell2txt(cells[j], sst, styles);
      }
    }
    for(i=0; i<data.length; i++) {
      for(j=0; j<data[i].length; j++) {
        txt += data[i][j] ? ('"' + (data[i][j]).toString().replace(/"/g,'""') + '"') : "";
        if (j<data[i].length-1) {txt += ",";}
      }
      txt += "\n";
    }
    callback(txt);
  };
  fr.readAsArrayBuffer(file);
};
