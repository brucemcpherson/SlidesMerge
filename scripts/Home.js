/**
* sets up all listeners
* @constructor Home
*/

var Home = (function (ns) {
  'use strict';

  // The initialize function must be run to activate elements
  ns.init = function (reason) {

    // select a folder for the result
    DomUtils.elem("folder-button").addEventListener ('click', function (e) {
      
      UsePicker.folderDialog()
      .then (function (packet) {

        if (packet.id) {
          Client.settings.deck.presoFolderPacket = packet;
        }
        displayPacket("presentation-result", "Result folder" , packet , Client.getFolderLink (packet.id));
      });
    });
    
    // select a template as input
    DomUtils.elem("presentation-button").addEventListener ('click', function (e) {
      UsePicker.presentationDialog()
      .then (function (packet) {
        if (packet.id) {
          Client.settings.deck.templatePacket = packet;
        }
        displayPacket("presentation-thumbnail", "Presentation template" , packet,  Client.getSlideLink (packet.id));
      });
    });
    
    // start generating
    DomUtils.elem("start-button").addEventListener ('click', function (e) {
      DomUtils.elem("start-button").disabled = true;
      Client.makeDecks ().then(function () { DomUtils.elem("start-button").disabled = false; });
    });
    
    function sbutton () {
      DomUtils.elem ("start-button").disabled = !(Client.settings.deck.templatePacket.id && Client.settings.deck.presoFolderPacket.id );
    }
    
    function displayPacket(elem, text , packet, link) {
    
     // I've got a new presentation to be the template
        var de = DomUtils.elem (elem);
        de.innerHTML = text;
        if (packet.id) {
         
          de.innerHTML = "";
          var te = DomUtils.addElem(de, "div");
          DomUtils.addElem (te,"img","","vimg").src=packet.iconLink;
          var a = DomUtils.addElem (te,"a",packet.title,"padded");
          a.href=link;
          a.target ="_blank";
          if (packet.thumbnail) {
            var sp = DomUtils.addElem (te,"div");
            DomUtils.addElem (sp,"img").src=packet.thumbnail;  
          }

        }
        sbutton();
    
    }
  };
  
  return ns;
  
})(Home || {});
