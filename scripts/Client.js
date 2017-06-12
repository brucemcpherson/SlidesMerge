/**
 * manage client side activity
 * @namespace Client
 */
var Client = (function (ns) {

  ns.settings = {
    pollTime: 10000,
    deck:{
     templatePacket: {}, //stuff about the template
     sheetPacket: {} ,  // stuff about the sheet if null, use the active sheet
     presoFolderPacket: {}, // stuff about the output folder
     globalsPacket: {},  // stuff about the globals variable source
     scopePacket: {},  // stuff about the sheets in scope
     options: {
      type: "multiple",    // multiple , single
      multiSuffix: "", // used with multiple, to append to deck name, gets data from sheet row variable
      startRow: 1,       // where to start
      finishRow: 3,       // where to finish
      masters:[],       // an array of slide positions (starting at 1 to not duplicate)
      nameBase:"slidesmerge-results"
     }
    },
    sheets:{},
    data:{}
  };
  
  // keep checking what sheets are in the book
  // and update drop downs if things have changed
  
  // I was concerned about memory leaks here (see this node issue .. https://github.com/nodejs/node/issues/6673)
  // however these would only occur if return loop() rather than just loop()
  ns.looping = function () {
    
    loopSheetManifest();
    
    function loopSheetManifest() {
      Promise.all ([ns.getSheetsInBook(), Provoke.loiter (ns.settings.pollTime)])
      .then (function () {
        loopSheetManifest();
      });
    }
    
  };

  ns.getSlideLink = function (id) {
    return "https://docs.google.com/presentation/d/" + id + "/edit";
  };
 
  ns.getFolderLink = function (id) {
    return "https://drive.google.com/drive/folders/" + id;
  };
  
  function showResultLinks (show,result) {
    // start by hiding them both
    DomUtils.hide ("slide-result-row", true);
    DomUtils.hide ("slide-folder-row" , true); 
    
              
   // links to final result .. 
    if (result) {
      DomUtils.elem ("slide-result-link").href = ns.getSlideLink (result.copies[0]);
      DomUtils.elem ("slide-result-link").innerHTML = Client.settings.deck.options.nameBase;
      DomUtils.elem ("slide-folder-link").href=ns.getFolderLink (result.folder);
      DomUtils.elem ("slide-folder-link").innerHTML =   Client.settings.deck.presoFolderPacket.title;
    }

    if (show) {
      DomUtils.hide (ns.settings.deck.options.type === "multiple" ? "slide-folder-row" : "slide-result-row" , false);             
    };
  };
  
 
  // start making decks
  ns.makeDecks = function () {
    
    spinCursor();
    // pick up any settings from the settings control elements
    ns.mapSettings ();
    
    // hide the results rows
    showResultLinks (false);
    
    return Provoke.run ('Server', 'start', ns.settings.deck)
    .then (function (result) {
      
      showResultLinks (true, result);
      App.toast ("deck created", "done");
      resetCursor();
    })
    ['catch'](function (err) {
      App.showNotification ("deck creation error", err);
      resetCursor();
    });
    
    
  };

  // adapt the selection for multi suffix
  ns.adaptMultiSuffix = function () {
    var er = Ui.settings.elementer;
    var current = er.getCurrent();
    var controls = er.getElements().controls;

    // get the current data sheet values
    return Provoke.run ("Server" , "getHeadings" , current.data) 
    .then (function (result) {
      sadapt ('multiSuffix',true,result);
    })
    ['catch'](function (err) {
      App.showNotification ("Getting data",err);
    });
  };
  
  // adapt the things that use sheets as a drop down
  ns.adaptSheets = function () {
    
    // these are the data variables, if there's a change then we need to update the potential multisuffix
    var names = ns.settings.sheets.sheets.map (function (d) { return d.name; });
    if (sadapt ('data' , false, names ,ns.settings.sheets.active) ){
       ns.adaptMultiSuffix(); 
    }
    
    // this is the globals variables 
    sadapt ('globals' , true, names);

  };
  
  function sadapt (key, includeBlanks, list, def) {
    var er = Ui.settings.elementer;
    var current = er.getCurrent();
    var controls = er.getElements().controls;
    var control = controls[key];
    
    // if we already have a selection keep it
    var original = current[key];;
    var options = DomUtils.getOptions (control);
    var exop = includeBlanks ? [""].concat (list) : list;
    
    // if options have changed, need to reset them
    if (JSON.stringify (options) !== JSON.stringify( exop)) {
      DomUtils.changeOptions (control , exop,  original);
    }
    
    // need to change the selected value
    if (!original || list.indexOf (original) === -1) {
      current[key] = def || "" ;
      er.applySettings (current); 
      control.value = current[key];
    }
    
    return original !== control.value;
    
  }
  
  // map settings needed for deck
  ns.mapSettings = function () {
    var sd = ns.settings.deck;
    var so = sd.options;
    var el = Ui.settings.elementer.getCurrent();
    
    Object.keys(so).forEach (function (k) {
      if (el.hasOwnProperty (k)) {
        so[k] = Utils.isNumeric(el[k]) ? parseFloat (el[k]) :el[k];
      }
      else {
        so[k] = "";
      }
    });
  
    // where to get the variables from
    // sheetId can also be set here, but this assumes we'll use the active sheet id
    ns.settings.deck.sheetPacket.sheetName = el.data;
    ns.settings.deck.globalsPacket.sheetName = el.globals;

    // stuff in scope (could pick this up server side, but in the future there may be a client input)
    ns.settings.deck.scopePacket = ns.settings.sheets;
  };
  
  // bring up the UI
  ns.init = function () {
    Ui.doElementer();
    Client.looping();
    resetCursor();
  };
  
  // get sheets that exist
  ns.getSheetsInBook = function () {
  
    return Provoke.run ("Server", "getSheetsInBook")
    .then (function (sheets) {
      ns.settings.sheets = sheets;
      ns.adaptSheets();
    })
    ['catch'](function (err) {
      App.showNotification  ("Error getting sheets", err); 
    });
    
  };
  
  function resetCursor() {
    DomUtils.hide ('spinner',true);
  }
  function spinCursor() {
    DomUtils.hide ('spinner',false);
  }
  
  return ns;
})(Client || {});