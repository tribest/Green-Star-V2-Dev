<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
     "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
	 
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
 <title>Greenstar, the original twin gear juicer!</title>
 <link rel="stylesheet" type="text/css" href="styles.css" />
 <link rel="stylesheet" type="text/css" href="transmenu.css" />
 <script type="text/javascript" src="transmenu.js"></script>
 <script language="javascript">
		function init() {
			//==========================================================================================
			// if supported, initialize TransMenus
			//==========================================================================================
			// Check isSupported() so that menus aren't accidentally sent to non-supporting browsers.
			// This is better than server-side checking because it will also catch browsers which would
			// normally support the menus but have javascript disabled.
			//
			// If supported, call initialize() and then hook whatever image rollover code you need to do
			// to the .onactivate and .ondeactivate events for each menu.
			//==========================================================================================
			if (TransMenu.isSupported()) {
				TransMenu.initialize();

				// hook all the highlight swapping of the main toolbar to menu activation/deactivation
				// instead of simple rollover to get the effect where the button stays hightlit until
				// the menu is closed.
				menu1.onactivate = function() { document.getElementById("products").className = "hover"; };
				menu1.ondeactivate = function() { document.getElementById("products").className = ""; };

				menu2.onactivate = function() { document.getElementById("onlinestore").className = "hover"; };
				menu2.ondeactivate = function() { document.getElementById("onlinestore").className = ""; };

				menu3.onactivate = function() { document.getElementById("moreinformation").className = "hover"; };
				menu3.ondeactivate = function() { document.getElementById("moreinformation").className = ""; };

				menu4.onactivate = function() { document.getElementById("affiliates").className = "hover"; };
				menu4.ondeactivate = function() { document.getElementById("affiliates").className = ""; };

				menu5.onactivate = function() { document.getElementById("support").className = "hover"; };
				menu5.ondeactivate = function() { document.getElementById("support").className = ""; };

				document.getElementById("umbria").onmouseover = function() {
					ms.hideCurrent();
					this.className = "hover";
				}

				document.getElementById("umbria").onmouseout = function() { this.className = ""; }
			}
		}
 </script>
 <meta http-equiv="imagetoolbar" content="no" />
</head>

<body onload="init()">
	<div id="parent">
		<h1 id="header">Greenstar, the original twin gear juicer</h1>
		<div id="content">
			<ul id="menu">
				<li id="products"><a href="#"></a></li>
				<li id="onlinestore"><a href="#"></a></li>
				<li id="moreinformation"><a href="#"></a></li>
				<li id="affiliates"><a href="#"></a></li>
				<li id="support"><a href="#"></a></li>
			</ul>
			
			
			<div id="home_left">
				<h3>The GreenStar GS-2000 <br />Complete Juicer and Food Processor</h3>
				<p>Green Star machines are reputably one of the world's best juice extractors and most efficient food processors. Using Exclusive Heavy Duty Twin Gear technology with a low 110 rpm, superior to those of any competitor, Green Star machines are your best choice. There is no need to look elsewhere. With their ease of use and simple clean up, Green Star Juice Extractors and Food Processors are award-winning machines of unsurpassed quality.</p>
				<p>It can also process whole foods, making favorites like almond and date pie crusts, pates from nuts and vegetables, sauces from a vast array of ingredients, baby foods, frozen fruit desserts and more. The choice is yours.</p>
				<img src="imgs/gs-side.png" alt="GreenStar 2000" class="left" />
				<fieldset id="gsparts">
					<legend>GS-2000 Parts List</legend>
					<ul id="gspartslist">
						<li class="grey"><a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GS028&Product_Count=8&Category_Code=JCJPGS">Fine Screen</a>
                                   		<li class="grey"><a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GS025G&Product_Count=5&Category_Code=JCACCGS">Glass Pitcher</a>
                              			<li class="grey"><a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GP030">Homogenizing Blank</a>
                                  		<li class="grey"><a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GS036&Product_Count=7&Category_Code=JCACCGS">Wooden Plunger</a>
                                  		<li class="grey"><a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GS038A&Product_Count=4&Category_Code=JCACCGS">Cleaning Brush</a>
                                  		<div class="clear" style="height: 5px; background: #fff;"></div>
						<li><a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GS005DT&Product_Count=0&Category_Code=JCACCGS">Drip Tray</a></li>
						<li><a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GP010&Product_Count=6&Category_Code=JCACCGS">Strainer</a></li>
						<li><a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GS029&Product_Count=7&Category_Code=JCJPGS">Coarse Screen</a></li>     	
					</ul>
				</fieldset>
				<div class="clear"></div>
				
				<img src="imgs/greybox.png" alt="grey = " class="left" style="margin-top: 7px;"/>
				<p style="text-indent: 0;">Items that are marked in grey are included with all GreenStar Models</p>
				
				<img src="imgs/greenbox.png" alt="green = " class="left" style="margin-top: 7px;" />
				<p style="text-indent: 0;">Items that are marked in green are specific to the GS-2000 Package</p>
	                        
	                    	<p>The Green Star GS-2000 Complete Juicer and Food Processor comes with all the parts of the GS-1000 plus the Coarse Screen, Strainer and Drip Tray juicing accessories. This is a great set for even more juicing options.</p>
				
				<div style="height: 70px; width: 330px; margin: 0 auto;">
                                	<a href="imgs/gs-side-1.png"><img src="imgs/gs-side2thumb.png" alt="" border="0" class="left" /></a>
                                	<a href="imgs/gs-front.png"><img src="imgs/gs-frontthumb.png" alt="" border="0" class="left" /></a>
                                	<a href="imgs/Twin-Gear.png"><img src="imgs/twingears.png" alt="" border="0" class="left" /></a>
                                	<a href="imgs/GreenStar-w_fruit.png"><img src="imgs/gs-fruitthumb.png" alt="" border="0" class="left" /></a>
                                	<p style="text-indent: 0; text-align: center;">Click on the above images for a larger viewing</p>
				</div>
				
				<div class="clear" style="height: 50px;"></div>
			</div>
			
			
			<div id="home_right">
				<h2 id="awards">awards</h2>
					<a href="#"><img src="imgs/stiftung.png" alt="stiftung warentest" border="0" id="stiftung" /></a>
					<a href="#"><img src="imgs/aliveaward.png" alt="alive award of excellence" border="0" id="alive" /></a>
					<a href="#"><img src="imgs/happyjuicer.png" alt="happy juicer choice awards" border="0" id="happyjuicer" /></a>
				<h2 id="recipes">recipes</h2>
					<img src="imgs/juice.png" alt="" style="margin-left: 20px;" />
						<%
							dim dsnr,conr,sqlr,rsr,numr
							numr = day(date)
							dsnr = "DBQ=" & Server.Mappath("databases/greenpower.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};"
					       		set conr=server.createobject("ADODB.connection")
							conr.open dsnr
							sqlr = "SELECT * FROM recipes"
							Set rsr = conr.Execute(sqlr)
							
					       		WHILE NOT rsr.EOF
								IF rsr("ID")=CINT(numr) THEN
							       		Response.Write "<h5><a href=recipe.asp?id=" & rsr("ID") & ">" & rsr("title") & "</a></h5>" & VbNewLine
									Response.Write "<p id=""recipe"">" & rsr("caption") & "</p>" & VbNewLine
						       		END IF
								rsr.MoveNext
							WEND
							rsr.Close
					       		conr.Close
						%>
					<div class="clear" style="height: 5px;"></div>
					<h5 style="color: #147A16;">Living With Green Power</h5>
					<a href="http://www.tribestlife.com/merchant2/merchant.mv?Screen=PROD&Store_Code=TL&Product_Code=GPBEM04&Product_Count=2&Category_Code=RB"><img src="imgs/greenpowerbook.png" alt="" border="0" style="margin-left: 20px; float: left;" /><p id="greenbook">The above recipe is brought to you by the courtesy of the copyrighter of "Living with Green Power" book. Click here if you want to pick up a copy.</a></p> 
					<div class="clear" style="height: 10px;"></div>	
				<h2 id="articles">articles</h2>
					<ul id="sidearticles">
						<%      
							dim dsn,con,sql,rs,num
			       		       		num = 0
			       		       		dsn="PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & Server.Mappath("databases/greenpower.mdb") & ";" & "JET OLEDB:Database Password=greenpower"
					       		set con=server.createobject("ADODB.connection")
					       		con.open dsn
					       		sql = "SELECT * FROM news WHERE ((Public=true)) ORDER BY ID DESC"
			       		       		Set rs = con.Execute(sql)
			       			  
							dim newNum
							newNum = 0
		                                        
							rs.MoveNext
							rs.MoveNext
							rs.MoveNext
							
			       				while newNum < 5
								newNum = newNum + 1
								if newNum = 1 or newNum = 3 or newNum = 5 then 
					       		       		color = " class=""scrow"" "
				       				else
							       		color = " "
								end if
				       		       		dim textd
				       		       		textd=Left(rs("Title"), 25)
				       		       		Response.Write "<li" & color & "><a href=news.asp?ID=" & rs("ID") & ">" & textd & "..." & "</a></li>" & VbNewLine
				       		       	
				       		       		rs.MoveNext
							wend
					       		rs.Close
			       		       		con.Close
		       				%>
					</ul>
			</div>
			
			<div class="clear"></div>
			
			<div id="footer">           
				<p style="margin; 0; text-align: center; background: #8BD15D; font-size: .8em; color: #fff;">
				© 2002-2004
				Tribest Corporation | 
		       		<a href="http://www.tribest.com">www.tribest.com</a> <br />
				(888) 254-7336
				(U.S. toll-free) | 

				(562) 623-7150
				(Outside the U.S.)</p>
			</div>
			<div class="clear"></div>
			
		</div>	
	</div>
	
		<script type="text/javascript" src="rendermenu.js"></script>
	
</body>
</html>