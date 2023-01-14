// -----------------------------------------
// Script permettant d'ajouter un lien hypertexte sur la sélection
// tout en appliquant sur style de caractères "Lien hypertexte"
// v1.1
// script rédigé par Emmanuel Côtez (emmanuel.cotez@mnhn.fr)
// -----------------------------------------

function char_style_exist(style_name) {
	if(app.activeDocument.characterStyles.item (style_name) == null) { 
		app.activeDocument.characterStyles.add({name:style_name});
	}	
	return app.activeDocument.characterStyles.item(style_name);
}
 
	function GetClipboard(){  
	    var clipboard;  
	    if(File.fs == "Macintosh"){  
	        var script = 'tell application "Finder"\nset clip to the clipboard\nend tell\nreturn clip';  
	        clipboard = app.doScript (script,ScriptLanguage.APPLESCRIPT_LANGUAGE);  
	    } else {  
	        var script = 'Set objHTML = CreateObject("htmlfile")\r'+  
	        'returnValue = objHTML.ParentWindow.ClipboardData.GetData("text")';  
	        clipboard = app.doScript(script,ScriptLanguage.VISUAL_BASIC);  
	    }  
	    
	    // DOI: 10.5852/fffxxsdf
	    var doi = clipboard.split("10.");
	    doi[1] = doi[1].replace("\t"," ") ;
	    doi[1] = doi[1].replace("\n","") ;

		// test du point à la fin du DOI (à supprimer le cas échéant)
	    doi = doi[1].split(" ");

		if (doi[0].charAt(doi[0].length-1)==".") { 
			doi[0] = doi[0].substring(0,doi[0].length-1) ;
			}	    

		return doi[0] ;
	}

	function generate() {
        return Math.floor((1 + Math.random()) * 0x10000).toString(16).substring(1);
        }

    var continuer=true;

	try { var mySel = app.selection[0]; 
	
  	var myLink = "https://doi.org/10." + GetClipboard() ;

  	mySel.contents = myLink ; 

  } catch (e) { 
  	alert ("No DOI in the clipboard; please copy a DOI first (any format of DOI)"); 
  	continuer=false;
   }

if (continuer==true) {
	
	// création du style de caractères "Lien hypertexte" s'il n'existe pas
  myStyle = char_style_exist("Lien hypertexte") ;

  app.findGrepPreferences=app.changeGrepPreferences=null;
  app.findGrepPreferences.findWhat=".+";
  app.changeGrepPreferences.changeTo="$0" ;

  myFinds = mySel.findGrep() ;  

    var count=0;
    if (mySel == null || mySel.contents.length == 0 || myLink==false ) {
    		        if (myLink != false) { alert ("No active selection; please select some text first"); }
    		    } else {
        
    	app.findGrepPreferences=app.changeGrepPreferences=null;
        app.findGrepPreferences.findWhat=".+";

        myFinds = mySel.findGrep() ;

        myLim = myFinds.length;
        for (var j=0; myLim > j; j++) {
        
        var name = myFinds[j].texts[0].contents ;        
        var MyLinkName = "DOI_"+myLink+"_"+generate() ;

	    try { 
	        var myHyperlinkURLDestination = app.activeDocument.hyperlinkURLDestinations.add(myLink) ; 
	    } catch (e) {
	        myHyperlinkURLDestination = myDoc.hyperlinkURLDestinations.itemByName(myLink);
	        if (myHyperlinkURLDestination == undefined) {
	            alert("problème avec une HyperDestination: " + liens[i].contents + "\nSupprimer la destination via le panneau HyperLines/Options de cible d'hyperlien");
	            }
	        }

	  try { 
	  	var myHyperlinkSource = app.activeDocument.hyperlinkTextSources.add(myFinds[j].texts[0]) ; 
		var myHyperlink = app.activeDocument.hyperlinks.add(myHyperlinkSource, myHyperlinkURLDestination, {name: MyLinkName+(count)}) ; 
	  }
	  catch (e) {}        
	  
       }
        count++;  
        }   
  app.changeGrepPreferences.appliedCharacterStyle="Lien hypertexte";
  mySel.changeGrep() ;
}