<div align="center">

## Number to Word Convertion


</div>

### Description

Converts Number into words up to 999 trillion. It is a simple short code.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Opal Raj Ghimire](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/opal-raj-ghimire.md)
**Level**          |Intermediate
**User Rating**    |3.9 (27 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/opal-raj-ghimire-number-to-word-convertion__1-55396/archive/master.zip)





### Source Code

<code>
<p>Number to words, fairly small code<br>
<font size="4">Opal R. Ghimire</font>, Kathmandu, email:buna48@hotmail.com<br>
 </p>
<hr>
<p>
<br>
<font color="#0000FF">'THIS FUNCTION CONVERTS 1 TO 999 TRILLION INTO WORDS</font><font color="#008000"><br>
</font><font color="#FF0000">Public Function</font> ToWords(Num
<font color="#FF0000">As String</font>) <font color="#FF0000">As String</font><br>
<font color="#FF0000">Dim</font> sFormated <font color="#FF0000">As String</font>, Unit
<font color="#FF0000">As String</font>, Ans(5) <font color="#FF0000">As String</font><br>
<font color="#FF0000">Dim</font> K <font color="#FF0000">As Integer</font>, K1
<font color="#FF0000">As Integer</font><br>
Ans(0) = "trillion ": Ans(1) = "billion ": Ans(2) = "million "<br>
Ans(3) =
"thousand ": Ans(4) = ""<br>
sFormated = Format(Num, "000000000000000.00")<br>
<font color="#FF0000">For</font> K = 1 <font color="#FF0000">To</font> 13
<font color="#FF0000">Step</font> 3<br>
   
Unit = Mid$(sFormated, K, 3)<br>
<font color="#FF0000">    If</font> Val(Unit) > 0 <font color="#FF0000">Then</font> ToWords = ToWords + ToNum(Unit) + Ans(K1)<br>
   
K1 = K1 + 1<br>
<font color="#FF0000">Next</font><br>
<font color="#0000FF">'HANDLES DECIMAL PARTS (IF ANY)</font><br>
<font color="#FF0000">If </font>Val(Num) - Int(Num) <> 0
<font color="#FF0000">Then</font> ToWords = ToWords + "and " +
Right$(sFormated, 2) +
"/100"<br>
<font color="#FF0000">End Function</font><br>
<br>
<br>
<font color="#0000FF">'THIS FUNCTION CONVERTS 1 TO 999 INTO WORDS</font><br>
<font color="#FF0000">Public Function</font> ToNum(Num <font color="#FF0000">As String</font>)
<font color="#FF0000">As String</font><br>
<font color="#FF0000">Dim</font> N(19) <font color="#FF0000">As String</font>, NN(8)
<font color="#FF0000">As String</font>, Formated <font color="#FF0000">As String</font><br>
<font color="#FF0000">Dim</font> Hun <font color="#FF0000">As Integer</font>, Tens
<font color="#FF0000">As Integer</font><br>
N(0) = "": N(1) = "one": N(2) = "two": N(3) = "three": N(4) = "four": N(5) =
"five": N(6) = "six": N(7) = "seven": N(8) = "eight": N(9) = "nine": N(10) =
"ten": N(11) = "eleven"<br>
N(12) = "twelve": N(13) = "thirteen": N(14) = "fourteen": N(15) = "fifteen":
N(16) = "sixteen": N(17) = "seventeen": N(18) = "eighteen": N(19) = "nineteen"<br>
NN(0) = "twenty": NN(1) = "thirty": NN(2) = "forty": NN(3) = "fifty": NN(4) =
"sixty": NN(5) = "seventy": NN(6) = "eighty": NN(7) = "ninety"<br>
Formated = Format(Num, "000.00")<br>
Hun = Mid$(Formated, 1, 1)<br>
Tens = Mid$(Formated, 2, 2)<br>
<br>
<font color="#FF0000">If</font> Hun <> 0 <font color="#FF0000">Then</font> ToNum = N(Hun) + " hundred "<br>
<font color="#FF0000">If</font> Tens <> 0 <font color="#FF0000">Then</font><br>
<font color="#FF0000">    If</font> Tens < 20 <font color="#FF0000">Then</font><br>
       
ToNum = ToNum + N(Tens) + " "<br>
<font color="#FF0000">    Else</font> '>20<br>
       
ToNum = ToNum + NN(Mid(Tens, 1, 1) - 2) + " " + N(Mid(Tens, 2, 1)) + " "<br>
<font color="#FF0000">    End If</font><br>
<font color="#FF0000">End If</font> <font color="#0000FF">'Tens <> 0</font><br>
<font color="#FF0000">End Function</font><br>
</code>

