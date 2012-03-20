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
				<h3>Differences Between Gold and GreenStar</h3>
				 
				<table class="modelcomparison">
					<tr class="mc_head">
						<td class="mc_right">Features</td>
						<td>GP-E1503</td>
						<td>GS-1000</td>
						<td>GS-2000</td>
						<td>GS-3000</td>
					</tr>
					
					<tr>
						<td class="mc_right">Juicing System</td>
						<td>Exclusive Heavy Duty Twin Gear</td>
						<td>Exclusive Heavy Duty Twin Gear</td>
						<td>Exclusive Heavy Duty Twin Gear</td>
						<td>Exclusive Heavy Duty Twin Gear</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Magnetic & bioceramic technology to extract more minerals</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr>
						<td class="mc_right">Pocket recess of each tooth of gear for easier juicing</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Built in mechanism to cut stringy fibers</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr>
						<td class="mc_right">Fine screen for vegetable juicing</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Coarse screen for pulpy juice</td>
						<td>Yes</td>
						<td>Optional</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr>
						<td class="mc_right">Homogenizing blank for making nut butters, baby foods, and fruit sorbets</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Pulp adjusting knob to control pressure on the pulp to produce maximum juice quantity</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr>
						<td class="mc_right">Plastic plunger</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Wooden plunger</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr>
						<td class="mc_right">Juice pitcher</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Carrying handle</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr>
						<td class="mc_right">Electrical cord storage compartment</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
						<td>Yes</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Pasta maker</td>
						<td>Yes</td>
						<td>Optional</td>
						<td>Optional</td>
						<td>Yes</td>
					</tr>
					
					<tr>
						<td class="mc_right">Mochi maker</td>
						<td>Yes</td>
						<td>Optional</td>
						<td>Optional</td>
						<td>Yes</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Pulp adjusting knob for soft fruit juicing</td>
						<td>Optional</td>
						<td>Optional</td>
						<td>Optional</td>
						<td>Optional</td>
					</tr>
					
					<tr>
						<td class="mc_right">Capacity of motor</td>
						<td>190 watts(1/4 hp)</td>
						<td>190 watts(1/4 hp)</td>
						<td>190 watts(1/4 hp)</td>
						<td>190 watts(1/4 hp)</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Squeezing HP</td>
						<td>4 hp</td>
						<td>4 hp</td>
						<td>4 hp</td>
						<td>4 hp</td>
					</tr>
					
					<tr>
						<td class="mc_right">Clearance between the two gears</td>
						<td>4/1000 inch</td>
						<td>4/1000 inch</td>
						<td>4/1000 inch</td>
						<td>4/1000 inch</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Switch button operation</td>
						<td>One touch operation</td>
						<td>One touch operation</td>
						<td>One touch operation</td>
						<td>One touch operation</td>
					</tr>
					
					<tr>
						<td class="mc_right">Warranty</td>
						<td>2 years</td>
						<td>5 years</td>
						<td>5 years</td>
						<td>5 years</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Machine & Accessories Weight</td>
						<td>26 lb</td>
						<td>21 lb</td>
						<td>22.5 lb</td>
						<td>23 lb</td>
					</tr>
					
					<tr>
						<td class="mc_right">Shipping Weight</td>
						<td>34 lb</td>
						<td>29 lb</td>
						<td>30.5 lb</td>
						<td>31 lb</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Size (W x H x D)</td>
						<td>22-1/4" x 10" x 15-1/2"</td>
						<td>20-3/14" x 10" x 15"</td>
						<td>20-3/14" x 10" x 15"</td>
						<td>20-3/14" x 10" x 15"</td>
					</tr>
				</table>
				
				<table class="modelcomparison" id="margintop">
					<tr class="mc_head">
						<td class="mc_right" style="width: 155px">Details</td>
						<td style="width: 80px;">GP-E1503</td>
						<td style="width: 80px;">GS-1000</td>
						<td style="width: 80px;">GS-2000</td>
						<td style="width: 80px;">GS-3000</td>
					</tr>
					
					<tr>
						<td class="mc_right">Suggested Retail Price</td>
						<td>$599</td>
						<td>$449</td>
						<td>$499</td>
						<td>$549</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Accessory, Outlet adjusting knob - soft</td>
						<td>$10.00</td>
						<td>$10.00</td>
						<td>$10.00</td>
						<td>$10.00</td>
					</tr>
					
					<tr>
						<td class="mc_right">Accessory, Rice cake blank (closed blank)</td>
						<td>Included</td>
						<td>$9.00</td>
						<td>$9.00</td>
						<td>Included</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Accessory, Rice cake guide</td>
						<td>Included</td>
						<td>$9.00</td>
						<td>$9.00</td>
						<td>Included</td>
					</tr>
					
					<tr>
						<td class="mc_right">Accessory, Pasta screen</td>
						<td>Included</td>
						<td>$25.00</td>
						<td>$25.00</td>
						<td>Included</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Accessory, Pasta screw</td>
						<td>Included</td>
						<td>$9.00</td>
						<td>$9.00</td>
						<td>Included</td>
					</tr>
					
					<tr>
						<td class="mc_right">Accessory, Pasta guide</td>
						<td>Included</td>
						<td>$9.00</td>
						<td>$9.00</td>
						<td>Included</td>
					</tr>
					
					<tr class="mc_alt_color">
						<td class="mc_right">Accessory, Drip Tray</td>
						<td>Included</td>
						<td>$9.00</td>
						<td>Included</td>
						<td>Included</td>
					</tr>
				</table>
				
				
				
				<div class="clear" style="height: 1px;"></div>
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