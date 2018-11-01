var Ui = (function (ns) {
  
  ns.settings = {};
  
  // set up the options manager
  ns.doElementer = function () {
    
    var setup = ns.setup();
    
    ns.settings.elementer = new Elementer()
    .setMain('')
    .setContainer('elementer-content')
    .setRoot('elementer-root')
    .setLayout(setup.layout)
    .setDetail(setup.detail)
    .build();
    
  };
  
  // the settings page 
  ns.setup = function() {

    return {
      detail: {
        
        masters: {
          template: "wideTemplate",
          label: "Master slide indices",
          icon: "speaker_notes_off",
          values:{
            value:""
          }
        },
        
        nameBase: {
          template: "wideTemplate",
          label: "Output base name",
          icon: "label",
          values:{
            value:"slides-merge-result"
          }
        },
        
        type: {
          template: "selectTemplate",
          label: "Output decks",
          icon: "merge_type",
          options:["single","multiple"],
          values:{
            value:"single"
          }
        },
        
        missingBehavior: {
          template: "selectTemplate",
          label: "Skip if data field(s) are null",
          icon: "featured_play_list",
          options:["never","any","all"],
          values:{
            value:"never"
          }
        },
        
        multiSuffix: {
          template: "selectTemplate",
          label: "Multi deck suffix",
          icon: "playlist_add",
          options:[""],
          values:{
            value:""
          }
        },

      
        startRow: {
          template: "numberTemplate",
          label: "Start row",
          icon: "first_page",
          values:{
            value:1
          },
          properties: {
            min:1
          }
        },
        
        finishRow: {
          template: "numberTemplate",
          label: "Finish row",
          icon: "last_page",
          values:{
            value:0
          },
          properties: {
            min:0
          }
        },
        
        data: {
          template: "selectTemplate",
          label: "Data sheet",
          icon: "swap_horiz",
          options:[],
          values:{
            value:""
          },
          on: {
            change: function (elementer , branch, ob , e){
              // need to change the drop down list for the suffix
              Client.adaptMultiSuffix();
            }
          }
        },
          
        globals: {
          template: "selectTemplate",
          label: "Global variable sheet",
          icon: "swap_vertical",
          options:[""],
          values:{
            value:""
          }
        },
        
        dataDivider: {
          template: "dividerTemplate",
          label: "Variables and data"
        },
        
        slideDivider: {
          template: "dividerTemplate",
          label: "Generated slides"
        }, 
        

        
      },
      
      layout: {
        settings: {
          prefix: "layout",
          root: "root"
        },
        
        pages: {
          root: {
            label:"Options",
            items:["slideDivider", "type","nameBase", "multiSuffix",
                   "masters","missingBehavior",
                   "dataDivider","data","startRow","finishRow","globals"],
            on: {
              exit: function (elementer, branch) {
               // actually this is only a one page settings, with no saving - so nothing to do
              }
            }
          }       
        }
      }
    }
  };
  
  
  return ns;
})({});
