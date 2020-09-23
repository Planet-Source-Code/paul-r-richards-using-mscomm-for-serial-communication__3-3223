<div align="center">

## Using MSComm for serial communication


</div>

### Description

How to use Microsofts Communication Control MSCOMM
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Paul R\. Richards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/paul-r-richards.md)
**Level**          |Beginner
**User Rating**    |3.5 (21 globes from 6 users)
**Compatibility**  |Microsoft Visual C\+\+
**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__3-3.md)
**World**          |[C / C\+\+](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/c-c.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/paul-r-richards-using-mscomm-for-serial-communication__3-3223/archive/master.zip)





### Source Code

```
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 9">
<meta name=Originator content="Microsoft Word 9">
<link rel=File-List href="./MSCOMM_files/filelist.xml">
<title>If you would like to use the Microsoft </title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
 <o:Template>Normal</o:Template>
 <o:LastAuthor>Paul Richards</o:LastAuthor>
 <o:Revision>2</o:Revision>
 <o:TotalTime>0</o:TotalTime>
 <o:Created>2002-01-31T02:59:00Z</o:Created>
 <o:LastSaved>2002-01-31T02:59:00Z</o:LastSaved>
 <o:Pages>2</o:Pages>
 <o:Words>273</o:Words>
 <o:Characters>1558</o:Characters>
 <o:Company> </o:Company>
 <o:Lines>12</o:Lines>
 <o:Paragraphs>3</o:Paragraphs>
 <o:CharactersWithSpaces>1913</o:CharactersWithSpaces>
 <o:Version>9.2720</o:Version>
 </o:DocumentProperties>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
@font-face
	{font-family:"MS Mincho";
	panose-1:0 0 0 0 0 0 0 0 0 0;
	mso-font-alt:"\FF2D\FF33 \660E\671D";
	mso-font-charset:128;
	mso-generic-font-family:roman;
	mso-font-format:other;
	mso-font-pitch:fixed;
	mso-font-signature:1 134676480 16 0 131072 0;}
@font-face
	{font-family:"\@MS Mincho";
	panose-1:0 0 0 0 0 0 0 0 0 0;
	mso-font-charset:128;
	mso-generic-font-family:roman;
	mso-font-format:other;
	mso-font-pitch:fixed;
	mso-font-signature:1 134676480 16 0 131072 0;}
 /* Style Definitions */
p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
p.MsoPlainText, li.MsoPlainText, div.MsoPlainText
	{margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Courier New";
	mso-fareast-font-family:"Times New Roman";}
@page Section1
	{size:8.5in 11.0in;
	margin:1.0in 65.95pt 1.0in 65.95pt;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
</head>
<body lang=EN-US style='tab-interval:.5in'>
<div class=Section1>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>If you
would like to use the Microsoft <o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>Communications
Control, here are the steps in VC:<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>1)
Create a dialog in VC++<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>2)
Project-&gt;Add to Project-&gt;Components and Controls<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>3) Find
&quot;Microsoft Communications Control&quot; click insert<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>4) Drag
the now visible phone icon onto your dialog<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>5)
Select phone icon and hit cntrl-w for class wizard<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>6)
Click on Member Variables and click on IDC_MSCOMM1<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>7) Add
a member variable name (m_Comm) for the control<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>8) You
can now open the serial port (1) using this code:<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>m_Comm.SetCommPort (1); // Com port 1<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>m_Comm.SetSettings
(&quot;38400,e,8,1&quot;);<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>m_Comm.SetInputMode (1); // Binary mode <o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>m_Comm.SetPortOpen (TRUE); // Open it<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>// Note: you may want to flush the buffer
here, by<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>// reading and throwing away the data <o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>// m_Comm.GetInput();<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>9) To
send data use this code:<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>CByteArray btArray;<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>btArray.Add (0x00); // Byte 1<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>btArray.Add (0x13); // Byte 2<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>COleVariant var(btArray);<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>m_Comm.SetOutput(var); // Send the data.<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'>10) To
read data use this code:<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>COleVariant myVar;<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>int hr;<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>long lLen = 0;<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>BYTE *pAccess;<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>char buffer[255];<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>myVar.Attach (m_Comm.GetInput());<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>// Get the length <o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>hr = SafeArrayGetUBound (myVar.parray, 1,
&amp;lLen);<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>if (hr == S_OK) <o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>{<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">       </span>lLen++; // upper bound is zero based
index<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">       </span>// lock array so you can access it.<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">       </span>hr = SafeArrayAccess
(myVar.parray,(void**)&amp;pAccess);<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">       </span>if (hr == S_OK) <o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">       </span>{<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">         </span>// Make a copy of the data <o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">         </span>// Note: Need to check that lLen is
&lt; buffer length<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">         </span>//<span style="mso-spacerun:
yes">       </span>so you don't overrun the buffer.<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">         </span>for (int i<span style="mso-spacerun:
yes">  </span>0; i &lt; lLen; i++)<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">            </span>buffer[i] = pAccess[i];<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><![if !supportEmptyParas]>&nbsp;<![endif]><o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">         </span>// unlock the data<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span><span style="mso-spacerun:
yes">    </span>SafeArrayUnaccessData (myVar.parray);<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">       </span>}<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>}<o:p></o:p></span></p>
<p class=MsoPlainText><span style='mso-fareast-font-family:"MS Mincho"'><span
style="mso-spacerun: yes">     </span>// COleVariant cleans itself up<o:p></o:p></span></p>
</div>
</body>
</html>
```

