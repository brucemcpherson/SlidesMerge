/** 
* all about using the Google Picker
* the API I dislike above all others
*/
var UsePicker = (function (ns) {

  /**
   * init the picker dialog
   */
  ns.settings = {
    dialog: {
      width:600,
      height:400
    },
    promises: {
    
    }
  };
  
  // Initialize the picker
  // get a token & make sure gapi gets loaded for later
  ns.init = function () {
    var sp = ns.settings.promises;
    
    // and load the picker
    sp.gapi =  new Promise (function (resolve, reject) {
        gapi.load('picker', {callback:resolve});
      });
  
    // and the developer key
    sp.devKey = Provoke.run ('PropertyStores', 'get' , 'script' , 'pickerDeveloperKey')
    .then (function (result) {
      // nothing to do for now
      return result;
    })
    ['catch'](function (err) {
      App.showNotification ("getting developerKey" , err);
    });
  };
  
  /**
  * do a folder dialog
  * @return {Promise}
  */
  ns.folderDialog = function () {
     var docsView = new google.picker.DocsView(google.picker.ViewId.FOLDERS)
     .setSelectFolderEnabled(true)
     .setIncludeFolders(true)
     .setOwnedByMe(true);

    return ns.dialog (docsView);
  };
  
  /**
  * do a presentations dialog
  * @return {Promise}
  */
  ns.presentationDialog = function () {
    
    var docsView = new google.picker.DocsView(google.picker.ViewId.PRESENTATIONS)
    .setIncludeFolders(true)

    return ns.dialog (docsView);
  };
  
  /**
  * @param {picker.view} viewId 
  * @return {Promise} 
  */
  ns.dialog = function (view) {
    
    var sp = ns.settings.promises, resolve, reject;
    
    // get a new token in case the old one expired
    sp.token = Provoke.run ('Server', 'getOAuthToken')
    .then (function (result) {
      // nothing to do for now
      return result;
    })
    ['catch'](function (err) {
      App.showNotification ("getting access token" , err);
    });
    
    // sort out the picker to return promises
    var pr = new Promise (function (res, rej) {
      resolve = res;
      reject = rej;
    });
    
    // when all is is ready, then bringup the picker
    return Promise.all ([sp.gapi, sp.token , sp.devKey])
    .then (function (r) {
      var token = r[1];
      var devKey = r[2];
      if (!token || !devKey) throw 'couldnt get credentials';

      // set up the picker
      var picker = new google.picker.PickerBuilder()
      .addView(view)
      .enableFeature(google.picker.Feature.NAV_HIDDEN)
      .hideTitleBar()
      .setOAuthToken(token)
      .setDeveloperKey(devKey)
      .setSize(ns.settings.dialog.width,ns.settings.dialog.height)
      .setCallback(function (data) {
        
        // called when something is picked
        // note that you get called on loaded to, so just resolve
        // when we have a picked or cancelled
        var action = data[google.picker.Response.ACTION];
        var picked = action == google.picker.Action.PICKED;
        var cancelled = action == google.picker.Action.CANCEL;
        var doc = picked && data[google.picker.Response.DOCUMENTS][0];
        var id = doc && doc[google.picker.Document.ID];
        var package = {
            action:action,
            picked:picked,
            doc:doc,
            id:id
          };
            
        // get a thumbnail and parents using drive API
        if (id) {
          
          var url = "https://www.googleapis.com/drive/v2/files/" + id;
          axios.get(url, {
            headers: {
              authorization: "Bearer " + token
            }
          })
          .then(function (response) {
            package.thumbnail = response.data.thumbnailLink;
            package.title = response.data.title;
            package.iconLink = response.data.iconLink;
            package.parents = response.data.parents;
           
            resolve (package);
          })
          ['catch'](function (error) {
            reject(error);
          });
        }
        else if (cancelled) {
          resolve (package);
        }
        

      })
      .setOrigin(google.script.host.origin)    
      .build();
      
      picker.setVisible(true);
      return pr;
    })
    ['catch'](function (err) {
      reject (err);
    });

  };

  return ns;

})({});









