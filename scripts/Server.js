/**
 * all communication with Slides API & sheets data
 * is done server side from this namespace
 */
var Server = (function(ns) {
  
  // local settings
  ns.settings = {
    chartPrefix:"chart",
    slidesPrefix:"slides",
    elementPrefixes:{
      "group":"elementGroup",
      "shape":"shape"
    },
    propertyKeys: { // not yet implemented
      resultsFolder:"slidesMerge_results_folder",
      templateDeck:"slidesMerge_template_deck",
      globalSheet:"slidesMerge_global_sheet"
    }
  };

  
  /**
  * make a map of every placeholder
  * this is used to optimize calls to slides api
  * and only call replacements for objects that actually contain placeholders
  */
  ns.getPlaceholderMap = function () {
    var st = ns.settings;
    var sp = st.params;
    var globals = sp.globals;
    var fiddler = st.package.fiddler;
    
    // from the data sheet
    var m =  fiddler.getHeaders()
    .reduce (function (p,c) {
      p[c] = {
        type:'data',
        appears:[]
      }
      return p;
    }, {});
    
    // add the placeholders mentioned in the global data 
    sp.placeholderMap = Object.keys(globals).reduce (function (p,c) {
      p[c] = {
        type:'global',
        appears:[]
      }
      return p;
    },m);


    
  };
  
  /**
  * make a map of how many times a placeholder appears on each slide
  * this is used to count placeholders per slide
  * later on, this wil be used to decide whether to keep a slide
  * if it meets the threshold for missing items
  */
  ns.makeAppearances = function () {  
    var st = ns.settings;
    var sp = st.params; 
    var fiddler = st.package.fiddler;
    
    // set up an appears map
    sp.appearsMap = Object.keys(sp.placeholderMap)
    .reduce(function (p,k) {
      sp.placeholderMap[k].appears.forEach (function (e) {
       
        p[e.slideId] = p[e.slideId] || {};
        p[e.slideId][k] = p[e.slideId][k] || {count:0};
        p[e.slideId][k].count++;
      });
      return p;
    }, {});
    
    // now blow out for each row of data
    sp.countMap = Object.keys(sp.appearsMap)
    .reduce (function (p,k) {
      // the current template slide
      var c = sp.appearsMap[k];
     
      // each row of data
      fiddler.getData().forEach (function (e,i) {
        
        // if its a single, then all the objects are in a single array element, indexed by name
        // if its a multi, then the number of array lements == no of rows in data
        var deck = sp.options.type === "multiple" ? i : 0 ;
        
        // count will be how many plaeholders on this page, observed will be updated for each one seen later
        var ob = {count:0 , observed:0 };
        
        // a multiple doesnt need a row number to id the object since it'll be in a seprate deck anyway
        var name = sp.options.type === "multiple" ? k : k+"_row_"+i   ;
        
        // this is a counter for each slide
        p[deck] = p[deck] || {};
        p[deck][name] =  ob;
        
        // sum the number of placeholders on this slide
        Object.keys(sp.placeholderMap).forEach (function (f) {
          if  (c[f]){
            ob.count += c[f].count;
          }
        });
      });
      return p;
    },[]);
   
    
  };
  
  /**
   * remove slides with missing placeholder data if required
   * because they might not meet the selected threshold for retention
   */
  ns.removeMissing = function () {
  
    var st = ns.settings;
    var sp = st.params; 
    var so = sp.options;
    

    var removals = sp.countMap.map (function (deck) {
      
      // keep all anyway
      if (so.missingBehavior === "never" ) return [];
      
      // filter out slides with missing placeholders
      return Object.keys(deck).filter(function (k) {
        // if we have them all, always keep
        var d = deck[k];
        if (d.count === d.observed || d.count === 0 ) return false;
        // but if we needed them all reject, or if we needed at least some reject if there's not any
        return so.missingBehavior === "any" ? true  : !d.observed;

      });

    });
    
    
    
    //now simply poke them on the delete requests if they are not already there
    removals.forEach (function (deck,i) {
      var dels = ns.settings.package.reqs[i].filter (function (e) {
        return e.hasOwnProperty ("deleteObject");
      })
      .map (function (e) {
        return e.deleteObject.objectId;
      });

                   
      deck.forEach (function (d) {
        
          if (dels.indexOf(d) === -1){
            
            ns.settings.package.reqs[i].push ({
              deleteObject: {
                objectId:d
              }
            });
          }

       
 
      });  
    });
  };
  
  
  // get the template contents
  // we want to make a map of where items appear to optimize the finl slides request
  ns.optimizePlaceholders = function () {
  
    var st = ns.settings;
    var sp = st.params;  
    var tep = sp.templatePacket;
    
    // fetch the template data - I only want the textruns
    // note that split placeholder will be ignored 
    var fields = "slides(objectId,pageElements(" + 
      "elementGroup(children(objectId,shape/text/textElements/textRun/content,table))," + 
      "objectId,shape/text/textElements/textRun/content,table" + 
     "))";
    
    // now we need to get the template data using the slides API
    var result =  Slides.Presentations.get(tep.id, {
      fields:fields
    });
   
    // the idea here is that we match up the slide numbers with where the placeholders appear
    result.slides.forEach (function (slide, idx) {
      
      // each pagelement within each slide
      slide.pageElements.forEach (function (pe , pi) {

        var shape = pe.shape;
        var table = pe.table;
        
        if (table) {
          table.tableRows.forEach (function (tr) {
            tr.tableCells.forEach (function (tc) {
              collectShapes ( tc && tc.text,pe ,idx , slide) ;
            });
          });
        }
        else {
          collectShapes ( pe.shape && shape.text,pe ,idx , slide) ;
        }        
        
      });
    });


    function collectShapes (text, pe , idx, slide) {
      if (text) {
        // look at all the text elements on a page
        text.textElements.forEach (function (te, ti) {
          
          // identified by the presence of a textRun.
          if (te.textRun) {
            
            // extract the content
            var content = te.textRun.content;
            
            
            // now we need to see if this is a place holder 
            Object.keys(sp.placeholderMap).forEach (function (k) {
              var rx = new RegExp("{{" + k + "}}");
              
              // if this matches, we've found a placeholder on this slide.
              if (content.match(rx)) {
                
                // placeholderMap is organized by placeholder key
                // and keeps a list of which pagelements in the template contain a given placeholder
                sp.placeholderMap[k].appears.push ({
                  pageElementId: pe.objectId,
                  slideIndex:idx,
                  slideId:slide.objectId
                });
                
              }
            });  
          }
        });
      }
    }    
   
  };

  // find out if there's any links to other decks
  // this whole section is unfinished and will not be implemented until
  // this isssue is resolved
  // https://issuetracker.google.com/issues/36761705
  ns.getFromOtherDecks = function () {
    var st = ns.settings;
    var sp = st.params;
    var se = st.elementPrefixes;
    sp.otherDecks = [];
    
    /*


    var fiddler = st.package.fiddler;
    var globals = sp.globals;
    
    // first the data
    sp.otherDecks = fiddler.getData()
    .reduce(function(p,c) {
      Object.keys(c).forEach (function (k) {
        addSlidesFetchPacket ( p , c[k] , k) ;
      });
      return p;
    } , {});
    
    // and the globals
    Object.keys (globals).reduce (function (p,c) {
      addSlidesFetchPacket ( p , globals[c] , c) ;
      return p;
    },sp.otherDecks);
    

    var fields = "slides(objectId,pageElements(objectId,transform," + Object.keys (se).map(function (k) { return se[k]; }).join (",") +"))";
    // now we need to get the data using the slides API
    var results = Object.keys(sp.otherDecks).map (function (e) {
      return Slides.Presentations.get(e, {
        fields:fields
      });
    });
    
    // now attach that to each of the required items
    Object.keys(sp.otherDecks).forEach (function (k, i) {
      var other = sp.otherDecks[k];
      var result = results[i];
      other.elements.forEach (function (d) {
        var slide = result.slides[d.slideIndex-1];
       
        if (!slide) throw 'slide ' + d.slideIndex + ' missing from deck for ' + d.value;
        
        // now find the element index that matches the type
        var elems = slide.pageElements.filter(function (e) {
          return e.hasOwnProperty (se[d.elementType]);
        });
        if (!elems[d.elementIndex-1]) throw 'element ' + d.elementType + "." + d.elementIndex + ' missing from deck for ' + d.value;
        d.elem = elems[d.elementIndex-1];
      });
      
    });
    
    // now make page element requests for each elementTODO!!!!!
    sp.pageElementRequests = Object.keys(sp.otherDecks).reduce (function (p,c) {
      var other = sp.otherDecks[c];
      other.elements.forEach (function (d,j) {
        p.push( {
          "pageObjectId": "x"+c+j,
          "size": {
            "width": {
              "magnitude": 3000000,
              "unit": 'EMU',
            },
            "height": {
              "magnitude": 3000000,
              "unit": 'EMU',
            },
          },
          "transform": {
            scaleX:0.6052, 
            scaleY:0.7583, 
            unit:"EMU",
            translateY:2791226.445, 
            translateX:1477066.6375000002
          }
        });
      });
      return p;
    },[]);
    

    function addSlidesFetchPacket (p, value , key) {
      // its a link to another deck
      // slides.id.sheetnumber.group.1
      var s = value.toString().split(".");
      if (s[0] === st.slidesPrefix && s.length > 1) {
        if (s.length !== 5 || s.some(function(e) { return !e;}) || !se.hasOwnProperty(s[3])) {
          throw 'invalid slides reference ' + value;
        }
        // get the id of the deck being referenced
        var id = s[1];
        if (!p.hasOwnProperty(id)) {
          p[id] = {
            id:id,
            elements:[]
          }
        };
        p[id].elements.push ({
          value:value,
          elementType:s[3],
          slideIndex:s[2],
          elementIndex:s[4],
          key:key
        });
      }
    }
    
    */
  }
  
  /**
   * gets the token for use client side picker use
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
   * set the parameters as default for next time
   */
  ns.setPosterity = function () {
    //not yet implemented - for the future
  };
  
  /**
   * entry point to create deck(s)
   */
  ns.start = function (params) {
  
    // get data parameters
    ns.settings.params = params;

    // open eveything
    ns.getEverything();
    
    // record properties for posterity against this document
    // will be implemented in a future version
    ns.setPosterity ();
    
    // make a map of all known placeholders
    ns.getPlaceholderMap();
    
    // optimize the placeholders to minimize requests later
    ns.optimizePlaceholders();
    
    // get a map of where placeholders appear
    ns.makeAppearances();
   
    // get info on any other decks that will be need to be referenced
    // this doesn't do anything for now until GAS issue is resolved
    ns.getFromOtherDecks();
  
    // generate the dup requests
    // returns [ each row [each slide] ]
    ns.createDupRequests();

    // copy files 
    ns.createFiles();
  
    // now apply subs
    ns.applySubs();
    
    // now deal wth deleting slides when placeholder values are missing
    ns.removeMissing ();
    
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
    var sp = ns.settings.params;
    
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
      
      // the slide ids associated with each row.
      var pobs =  dr[i].map (function (e) {
        return e.objectIds[e.objectId];
      });
      
      // do all the substitutions
      // this is a single deck, so the object will be named and in a single countmap element
      doSubs ( p , pobs, headers ,d, sp.countMap[0]);

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
    var sp = ns.settings.params;
    
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
      // this is a multi deck, so the object will be named and in a countmap  element that matches its row number
      doSubs ( p , pobs, headers,d,sp.countMap[i]);
      return p;
    });
    
  };
  
  /**
  * do the subs
  * @param {string[]} pobs an array of object ids, one for each row in the data
  */
  function doSubs (reqs,pobs,headers,dt,countMap) {
    
    var ss = ns.settings.package.ss;
    var sheet = ns.settings.package.sheet;
    var st = ns.settings;
    var sp = st.params; 
    
    
    function tweakPobs (placeholder, pobs) {
    
      var p = sp.placeholderMap[placeholder];
      if (!p) throw 'somethings gone wrong with optimization:Lost placeholder ' + placeholder ;
      
      // reduce the pobs to only contain slide objects that reference the placeholder
      // the pobs is the slideId + some row number
      return pobs.filter (function (d) {
        return p.appears.some (function (e) {
          return d.slice(0,e.slideId.length) === e.slideId;
        });
      });
    }
    
    // global subs - one for each known global
    Object.keys(ns.settings.params.globals).forEach (function (h) {
      var v = ns.settings.params.globals[h];
      var s = v.value;
      
      // tweak the list of slide objects to apply this to exclude those in which 
      // the placeholder doesn't appear
      var tweaked = tweakPobs (h, pobs);

      // chart sub
      chartImageSub (reqs,v,h,tweaked);
      chartSub (reqs,v,h,tweaked);
        
      // image sub
      imageSub (reqs ,  s, h ,tweaked);
        
      // text subs
      textSub (reqs , s , h ,tweaked,countMap);
    });
        
      
    // substitute values from data
    headers.forEach (function (h) {
      
      var v = resolveCharts (ss, dt[h].toString() , sheet)
      var s = dt[h].toString();
      
      // the placeholder doesn't appear
      var tweaked = tweakPobs (h, pobs);

      chartImageSub (reqs,v,h,tweaked);
      chartSub (reqs,v,h,tweaked);
      
      // image substitutions
      imageSub (reqs , s , h ,tweaked);
        
      // text substitutions
      textSub (reqs , s , h ,tweaked,countMap);
      
    });

    
  }
  
  
  
  // substitute global values
  function imageSub (reqs , text , field, pobs) {
    
    // image subs
    if (pobs.length && text.slice(0,4) === "http") {
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

  
  // substitute a chart as an image
  function chartImageSub (reqs , tob , field , pobs) {
    
    if (pobs.length && tob.type === ns.settings.chartPrefix) {
      
      
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
  
  // substitute as a linked chart
  function chartSub (reqs , tob , field , pobs) {
    
    if (pobs.length && tob.type === ns.settings.chartPrefix) {
      
      
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
  
  /**
  * do the text substitution
  * the list of objects has been optimized so that it only applies to slides that contain 
  * the placeholder currently being worked on
  * so if pobs.length ===0 then the placeholder doesn't appear so skip
  */
  function textSub (reqs , text , field , pobs, countMap) {
    
    var st = ns.settings;
    var sp = st.params; 
    
    // text substitution
    if (pobs.length) {
      
      // count observations on slide
      if (text !== "") {
        pobs.forEach (function (e) {
          countMap[e].observed ++;
        });
      }
      reqs.push ({ replaceAllText:{
        replaceText:text,
        pageObjectIds:pobs,
        containsText:{
          text:"{{" + field + "}}",
          matchCase:true
        }
      }});
      
    }
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
  // the dup requests contain a batch request list to duplicate - one for every row in the data
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

  // charts have a special syntax
  // pointing to potentially  different spreadsheet as the source
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
  



  return ns;
})(Server || {});
