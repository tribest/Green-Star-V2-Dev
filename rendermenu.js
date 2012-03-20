if (TransMenu.isSupported()) {

		//==================================================================================================
		// create a set of dropdowns
		//==================================================================================================
		// the first param should always be down, as it is here
		//
		// The second and third param are the top and left offset positions of the menus from their actuators
		// respectively. To make a menu appear a little to the left and bottom of an actuator, you could use
		// something like -5, 5
		//
		// The last parameter can be .topLeft, .bottomLeft, .topRight, or .bottomRight to inidicate the corner
		// of the actuator from which to measure the offset positions above. Here we are saying we want the 
		// menu to appear directly below the bottom left corner of the actuator
		//==================================================================================================
		var ms = new TransMenuSet(TransMenu.direction.down, 0, 0, TransMenu.reference.bottomLeft);

		//==================================================================================================
		// create a dropdown menu
		//==================================================================================================
		// the first parameter should be the HTML element which will act actuator for the menu
		//==================================================================================================
		var menu1 = ms.addMenu(document.getElementById("products"));
		menu1.addItem("GS-1000", "http://www.greenstar.com/"); 
		menu1.addItem("GS-2000", "gs2000.asp"); 
		menu1.addItem("GS-3000", "http://www.greenstar.com/"); 
		menu1.addItem("Gold GP-E1503", "http://www.greenstar.com/"); 
		
		
		//==================================================================================================
		//==================================================================================================
		var menu2 = ms.addMenu(document.getElementById("onlinestore"));
		menu2.addItem("Juicer Parts & <br />Accessories", "http://www.greenstar.com/");
		menu2.addItem("Tribestlife Store", "http://www.tribestlife.com/");

		//==================================================================================================
		//==================================================================================================
		var menu3 = ms.addMenu(document.getElementById("moreinformation"));
		menu3.addItem("Model Comparison Chart", "modelchart.asp");
		menu3.addItem("Why is it so Special?", "http://www.greenstar.com/");
		menu3.addItem("10 Ways to Maximize <br />Nutritional Value", "http://www.greenstar.com/");
		menu3.addItem("Twin Gear Technology", "http://www.greenstar.com/");
		menu3.addItem("Five Machines in One!", "http://www.greenstar.com/");
		menu3.addItem("Special Features of <br />Green Power", "http://www.greenstar.com/");
		menu3.addItem("Start a Healthy Day <br />with Green Power", "http://www.greenstar.com/");
		menu3.addItem("F. A. Q.", "http://www.greenstar.com/");
		menu3.addItem("Operation Manual", "http://www.greenstar.com/");
		menu3.addItem("Product Registration", "http://www.greenstar.com/");
		menu3.addItem("Find Retail Outlets", "http://www.greenstar.com/");
		menu3.addItem("Recent News Articles", "http://www.greenstar.com/");
		menu3.addItem("News Letters", "http://www.greenstar.com/");

		//==================================================================================================
		//==================================================================================================
		var menu4 = ms.addMenu(document.getElementById("affiliates"));
		menu4.addItem("SoloStar Juicers", "http://www.solostarjuicers.com"); 
		menu4.addItem("CitriStar Juicers", "http://www.citristar.com");
		menu4.addItem("Z-Star Juicer", "http://www.zstarjuicer.com");
		menu4.addItem("Personal Blender", "http://www.personalblender.com");
		menu4.addItem("Freshlife Sprouter", "http://www.freshlifesprouter.com");
		menu4.addItem("Purewise Distiller", "http://www.tribest.com/purewise.asp");
		menu4.addItem("Hawo's Grain Mills", "http://www.hawosmill.com");
		menu4.addItem("Wolfgang Mill", "http://www.wolfgangmill.com");   
		menu4.addItem("SilenX Computers", "http://www.silenx.com");
		
		//==================================================================================================
		//==================================================================================================
		var menu5 = ms.addMenu(document.getElementById("support"));
		menu5.addItem("(888) 254-7336", ""); 
		menu5.addItem("(562) 623-7150", "");
		menu5.addItem("About Us", "http://www.greenstar.com/");
		
		

		//==================================================================================================
		//==================================================================================================
		// write drop downs into page
		//==================================================================================================
		// this method writes all the HTML for the menus into the page with document.write(). It must be
		// called within the body of the HTML page.
		//==================================================================================================
		TransMenu.renderAll();
	}