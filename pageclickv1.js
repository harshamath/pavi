//Phantomjs script to get HTML of product page where products listed are more than 50 and paginated.

var args = require('system').args;
var page = require('webpage').create();

var url = args[1];
var plink = args[2];

page.onConsoleMessage = function(msg) {
    console.log(msg);
};

page.onAlert = function(msg) {
    console.log('alert!!> ' + msg);
};

page.onLoadStarted = function() {
    loadInProgress = true;
    console.log("load started");
};

page.onLoadFinished = function(status) {
    loadInProgress = false;
    if (status !== 'success') {
        console.log('Unable to access network');
        phantom.exit();
    } else {
        console.log("load finished");
    }
};

page.open(url, function(status) {	
    if(status === "success") {
        		
		setTimeout(function(){/* Function to Click on drop down to select 50 products to be displayed per page*/
			//page.render('screen1.png');
			page.evaluate(function(){
				var e = document.createEvent('MouseEvent');
				setTimeout(function(){
					e.initMouseEvent('click', true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
					document.querySelector("#ctl00_cphNestedMasterPage_gvItems_DXPagerTop_PSP_DXI2_T").dispatchEvent(e);
				},3000);
				//console.log("Select Clicked");
			});	
		
		
			setTimeout(function(){/*Function to Click on paginated links*/
				//page.render('screen2.png');
				page.evaluate(function(p){
					var e = document.createEvent('MouseEvent');
					e.initMouseEvent('click', true, true, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
					//document.querySelector("#ctl00_cphNestedMasterPage_gvItems_DXPagerTop > a:nth-child(4)").dispatchEvent(e);
					document.querySelector(p).dispatchEvent(e);
					//console.log("Page Clicked");
				}, plink);
				
				setTimeout(function(){/*Return HTML of product page*/
					//page.render('screen3.png');
					var html = page.evaluate(function() {
						return document.documentElement.outerHTML;
					});
					console.log(html);
					phantom.exit();
				},10000);
				
			},10000);
					
        },10000);
		
	}
	   
});