
var Server = (function(ns) {
  
  ns.settings = {
    chartPrefix:"chart"
  };

  /**
   * gets the token for use client side
   */
  ns.getOAuthToken = function () {
    return ScriptApp.getOAuthToken();
  };
  
  /**
   * get headings for a given sheet
   * @param {string} sheetName
   * @return {string[]} heading names
   */
  ns.getHeadings = function (sheetName) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  };
  /**
   * entry point to create deck(s)
   */
  ns.start = function (params) {
  
    // get data parameters
    ns.settings.params = params;

    // open eveything
    ns.getEverything();
  
    // generate the dup requests
    // returns [ each row [each slide] ]
    ns.createDupRequests();

    // copy files 
    ns.createFiles();
  
    // now apply subs
    ns.applySubs();
  
    // now apply   
    ns.execute ();
    
    // return the ids of what's been donw
    return {
      folder:ns.settings.package.presoFolder.getId(),
      copies:ns.settings.package.copies.map(function(d) {
        return d.getId();
      })};
    
  
  };
  
  // ready to go
  ns.execute = function () {
  
    return Server.settings.package.copies.map (function (d,i) {
      return Slides.Presentations.batchUpdate({'requests': ns.settings.package.reqs[i]}, d.getId());
    });
    

  };
  
  // get names of all sheets in workbook
  // also get all the charts int he book
  ns.getSheetsInBook = function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // we need to use the advanced sheets api to get the real chart ID
    var ch = Sheets.Spreadsheets.get(ss.getId(), {"fields": "sheets(charts/chartId,properties(sheetId,title))"});
    var sheets = ch.sheets.map (function (d) {
      return {
        id:d.properties.sheetId,
        name:d.properties.title,
        charts:d.charts
      }
    });
   

    return {
      active:ss.getActiveSheet().getName(),
      sheets:sheets
    }; 
  };
  
  // enhance duprequests with text subs
  ns.applySubs = function () {
    
   
    if (ns.settings.params.options.type === "multiple") {
      return applySubsMulti();
    }
    
    // this is the single version which is a lot more complicated
    var dr = ns.settings.package.dupRequests;
    var fiddler = ns.settings.package.fiddler;
    var headers = fiddler.getHeaders();
    
    // this is the sheet data
    var data = ns.settings.package.fiddler.getData();
    
    // this'll be used to put them in the correct position later
    var m = ns.settings.params.options.masters;
    var insertions = [], masters = [], aMasters = 
        m ? (Array.isArray(m) ? m : m.toString().split(",").map(function(d) { return parseInt(d,10);})) : [] ;
 
    
    ns.settings.package.reqs = data.map  (function (d,i) {
      
      // duplicate all slides in template & move to correct position
      var p = dr[i].reduce (function (t,c,j) {
        t.push ( {duplicateObject:c});
        
        // and do this later on once all the slides have been created
        // because dup puts them in a daft order
        insertions.push  ({
          updateSlidesPosition:{
            slideObjectIds : [c.objectIds[c.objectId]],
            insertionIndex: dr[i].length * i + j 
          }});
        
        // and we'll not need to be duplicating masters, except for the first one
        if (i && aMasters.indexOf(j+1) !== -1) {
          masters.push ({
            deleteObject:{
              objectId : c.objectIds[c.objectId]
            }});
        }
        
        return t;
      },[]);
      
      // common to all
      var pobs =  dr[i].map (function (e) {
        return e.objectIds[e.objectId];
      });
      
      // do all the substitutions
      doSubs ( p , pobs, headers ,d);

      return p;
    });
    

    // delete the original templates
    var dels = ns.settings.package.objectIds.map(function(e) {
      return {
        deleteObject: {
          objectId:e
        }
      } 
    });
    Array.prototype.push.apply (ns.settings.package.reqs, dels);

    // sort
    Array.prototype.push.apply (ns.settings.package.reqs, insertions);
    
    // delete any masters that dont need to be duplicated
    Array.prototype.push.apply (ns.settings.package.reqs, masters);

    // wrap in array to be compat with multi
    ns.settings.package.reqs =[ns.settings.package.reqs];
    
  };
  
  // for when we're creating multiple files.
  function applySubsMulti () {
    
    
    var dr = ns.settings.package.dupRequests;
    var fiddler = ns.settings.package.fiddler;
    var headers = fiddler.getHeaders();
   
    
    // this is the sheet data
    var data = ns.settings.package.fiddler.getData();
    
    // need a set of reqs for each file
    ns.settings.package.reqs = data.map(function (d,i) {

      var p = [];
      // common to all
      var pobs =  dr[i].map (function (e) {
        return e.objectId;
      });
      
      // do all the substitutions
      doSubs ( p , pobs, headers,d);
      return p;
    });
    
  };
  
  // do the subs
  function doSubs (reqs,pobs,headers,dt) {
    
    var ss = ns.settings.package.ss;
    var sheet = ns.settings.package.sheet;
    
    
    // global subs
    Object.keys(ns.settings.params.globals).forEach (function (h) {
      var v = ns.settings.params.globals[h];
      var s = v.value;
      
      // chart sub
      chartImageSub (reqs,v,h,pobs);
      chartSub (reqs,v,h,pobs);
        
      // image sub
      imageSub (reqs ,  s, h ,pobs);
        
      // text subs
      textSub (reqs , s , h ,pobs);
    });
        
      
    // substitute values from data
    headers.forEach (function (h) {
      
      var v = resolveCharts (ss, dt[h].toString() , sheet)
      var s = dt[h].toString();
      chartImageSub (reqs,v,h,pobs);
      chartSub (reqs,v,h,pobs);
      
      // image substitutions
      imageSub (reqs , s , h ,pobs);
        
      // text substitutions
      textSub (reqs , s , h ,pobs);
      
    });

    
  }
  
  
  
  // substitute global values
  function imageSub (reqs , text , field, pobs) {
    
    // image subs
    if (text.slice(0,4) === "http") {
      reqs.push ({ replaceAllShapesWithImage:{
        imageUrl:text,
        replaceMethod: 'CENTER_INSIDE',
        pageObjectIds:pobs,
        containsText:{
          text:"{{{" + field + "}}}",
          matchCase:true
        }
      }});
    }
    return reqs;
  }

  
  
   function chartImageSub (reqs , tob , field , pobs) {
     
     if (tob.type === ns.settings.chartPrefix) {

       
       // chart substitution - {{UNLINKED}}
       reqs.push ({ replaceAllShapesWithSheetsChart:{
         spreadsheetId:tob.id,
         chartId:tob.chartId,
         pageObjectIds:pobs,
         linkingMode:"NOT_LINKED_IMAGE",
         containsText:{
           text:"{{{" + field + "}}}",
           matchCase:true
         }
       }});

     }
  }
  
   function chartSub (reqs , tob , field , pobs) {
     
     if (tob.type === ns.settings.chartPrefix) {

       
      // chart substitution - {{LINKED}}
       reqs.push ({ replaceAllShapesWithSheetsChart:{
         spreadsheetId:tob.id,
         chartId:tob.chartId,
         pageObjectIds:pobs,
         linkingMode:"LINKED",
         containsText:{
           text:"{{" + field + "}}",
           matchCase:true
         }
       }});
     }
  }
  
  function textSub (reqs , text , field , pobs) {
    
    // text substitution
    reqs.push ({ replaceAllText:{
      replaceText:text,
      pageObjectIds:pobs,
      containsText:{
        text:"{{" + field + "}}",
        matchCase:true
      }
    }});
  }
  
  //
  // create the new files
  //
  ns.createFiles = function () {
    
    var sx = ns.settings.package;
    var so = ns.settings.params.options;
    var sp = ns.settings.params;
   
    // if its a single type, we only need to create one file
    if (so.type === "single") {
      sx.copies = [sx.templateFile.makeCopy(so.nameBase, sx.presoFolder)];
    }
    else if (so.type === "multiple") {
      sx.copies = sx.fiddler.getData().map (function(d,sindex) {
        var n =  (so.nameBase + "-" + (so.multiSuffix ? d[so.multiSuffix] : sindex + (so.startRow || 1)));
        return sx.templateFile.makeCopy(n, sx.presoFolder)
      });
    }
    else {
      throw 'unknown type ' + so.type;
    }

  };
  
  //
  // the dup requests contain a batch request to duplicate - one for every row in the data
  // 
  ns.createDupRequests = function () {
    
    var sk = ns.settings.package;
   
    sk.dupRequests = sk.fiddler.getData().map(function (d,i) {
      return  sk.objectIds.map (function (e) {
        var eid = {};
        eid[e] = e + "_row_" + i;
        return  { 
          objectId: e,
          objectIds: eid
        };
      });
    });
  };
  
  /**
  * create a data package of everything we'll need to make this happen
  */
  ns.getEverything = function() {
    
    // some short cuts
    var st = ns.settings;
    var sp = st.params;  
    var tep = sp.templatePacket;
    var shp = sp.sheetPacket;
    var prp = sp.presoFolderPacket;
    var stp = sp.globalsPacket;
    
    
    // open the sheet 
    var ss = shp && shp.sheetId ? SpreadsheetApp.openById(shp.sheetId) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw "could not open spreadsheet " + shp.sheetId;
    
    var sheet = shp && shp.sheetName ? ss.getSheetByName(shp.sheetName) : ss.getActiveSheet();
    if (!sheet) throw "could not open target sheet";

    // get the sheet data and filter out what we don't need
    var fiddler = new Fiddler()
    .setValues(sheet.getDataRange().getValues())
    .filterRows (function ( row, properties) {
      var rn = properties.rowOffset +1;
      return (!sp.options.startRow || rn  >= sp.options.startRow ) && (!sp.options.finishRow || rn <= sp.options.finishRow) ;
    });
   

    // get the globals variable packet if there is one
    // although the add-on only provides global variables in same spreadsheet, this will handle both for future
    if (stp && stp.sheetName) {
      var sss = stp && stp.sheetId ? SpreadsheetApp.openById(stp.sheetId) : ss;
      if (!sss) throw "could not open spreadsheet " + stp.sheetId;
      var stpSheet = sss.getSheetByName(stp.sheetName);
      if (!stpSheet) throw "could not open global variable sheet " + stp.sheetName;
      var stpFiddler = new Fiddler().setValues(stpSheet.getDataRange().getValues());
      var heads = stpFiddler.getHeaders();
      if (heads.indexOf("name") === -1 || heads.indexOf("value") === -1) {
        throw 'need name and value columns in global variable sheet';
      }
      // now we can set up the globals
      sp.globals = stpFiddler.getData()
      .reduce(function (p,c) { 
        // ignore rows with no reasonable name
        if (c.name) {
          if (p.hasOwnProperty(c.name)) {
            throw 'Duplicate name ' + c + ' in globals sheet';
          }
          p[c.name] = c.value;
        }
        return p;
      },{});

      // any chart stuff.. find any positional and rename to chart id
      Object.keys(sp.globals)
      .forEach(function(k) {
        sp.globals[k] = resolveCharts(ss, sp.globals[k]);
      });

      
    }
    else {
      sp.globals = {};
    }
    

    
    // get the template
    var template = DriveApp.getFileById(tep.id);
    if (!template) throw "could not open slides template " + sp.template;

    // and the folder - we'll duplicate the slide to there
    var folder = DriveApp.getFolderById(prp.id);
    if (!folder) throw "could not open output folder ";

    // and the deck
    var deck = Slides.Presentations.get(template.getId());
    if (!deck) throw "could not get slides";
   
    // and we're good to go
    st.package = {
      presoFolder: folder,
      fiddler: fiddler,
      sheet: sheet,
      ss: ss,
      templateFile: template,
      deck:deck,
      objectIds:getObjectIds(deck)
    };
    
    
  };

  function resolveCharts (ss, value, sheet) {
    var v = {};
    var st = ns.settings;
    var scp = st.params.scopePacket;
    v.value = value.toString();
    v.id = ss.getId();
    var match = v.value.match(/[^."']+|"([^"]*)"|'([^']*)'/g);
    // is this a chart alias?
    
    if (match && match[0] === st.chartPrefix) {
      
      // may do the same with things other than charts later
      v.type = match[0];
      
      // it's on the non default data sheet
      if (match.length === 3 ) {
        var cindex = match[2];
        var cs =  ss.getSheetByName(match[1]);
        if (!cs) throw  'cannot find chart alias sheet ' + v.value;
      }
      
      // its on the default data sheet
      else if (match.length === 2)  {
        var cindex = match[1];
        var cs = sheet;
      }
      
      // its a mess
      else {
        throw 'invalid chart alias ' + v.value;
      }
      
      // we'll need the chartId and sheetId later         
      v.sheetId = cs.getSheetId();
      v.chartIndex = Utils.isNumeric(cindex) ? parseInt(cindex) : 0;
      
      
      // we already have the in scope charts
      var charts = scp.sheets.filter (function (e) {
        return e.id === v.sheetId && e.charts && e.charts.length;
      })[0];
      
      // make sure there are some
      if (!charts) throw 'no charts found in sheet '+ cs.getName();
      
      // if there is no cindex, then we have to assume that we got the actual id
      // otherwise it's positional
      v.chartId = v.chartIndex ? 
        (charts.charts[v.chartIndex-1] ? charts.charts[v.chartIndex-1].chartId : "") : cindex;

      // if we didnt get it, its still a mess
      if (!v.chartId) throw 'coundnt find matching chart for ' + v.value;
      
    }
    return v;
  }
  
  /** 
  * this one will find all the objectids in a deck
  */
  function getObjectIds (preso) {
    
    // find the objectIds of interest on each slide
    return preso.slides.map(function(s) {
      return s.objectId;
    });
    
  }
  
  /**
   * server side util to find folder from path
   */
  ns.getDriveFolderFromPath = function (path) {
    return (path || "/").split("/").reduce ( function(prev,current) {
      if (prev && current) {
        var fldrs = prev.getFoldersByName(current);
        return fldrs.hasNext() ? fldrs.next() : null;
      }
      else { 
        return current ? null : prev; 
      }
    },DriveApp.getRootFolder()); 
}


  return ns;
})(Server || {});
