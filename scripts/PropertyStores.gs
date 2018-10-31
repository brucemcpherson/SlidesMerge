/**
 * use this to access property stores from client side
 * @namespace PropertyStores
 */
var PropertyStores = (function (ns) {
  
  ns.settings = {};
  
  function makeService (service) {
    
    return {
      store:service,
      get: function (key) {
        var r = this.store.getProperty (key);
        try {
          var ob = r ? JSON.parse(r) : null;
        }
        catch (err) {
          var ob = r;
        }
        return ob;
      },
      set: function (key , ob) {
        return this.store.setProperty (key , JSON.stringify(ob));
      }
    };
  }
  
  ns.init = function () {
    ns.settings.props = {
        script: makeService (PropertiesService.getScriptProperties()),
        doc: makeService (PropertiesService.getDocumentProperties()),
        user: makeService (PropertiesService.getUserProperties())
      };
    
  };

  
  /**
   * get an item from a given store
   * @param {string} store name (script|doc|user)
   * @return {*} the value
   */
  ns.get = function (store , key) {
    // if not already init - then do it now
    if (!ns.settings.props)ns.init();
    
    return ns.settings.props[store].get(key);
  };
  return ns;
})({});
