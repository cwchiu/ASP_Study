<%    
'*************************************   
'¨¾¤î¥~³¡´£¥æ   
'*************************************   
function ChkPost()    
  dim server_v1,server_v2   
  chkpost=false   
  server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))   
  server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))   
  If Mid(server_v1,8,Len(server_v2))<>server_v2 then   
    chkpost=False  
  else   
   chkpost=True  
  end If  
 end function   
  
'*************************************   
'IP¹LÂo   
'*************************************    
function MatchIP(IP)   
 on error resume next   
 MatchIP=false   
 Dim SIp,SplitIP   
 for each SIp in FilterIP   
    SIp=replace(SIp,"*","\d*")   
    SplitIP=split(SIp,".")   
    Dim re, strMatchs,strIP   
     Set re=new RegExp   
      re.IgnoreCase =True  
      re.Global=True  
      re.Pattern="("&SplitIP(0)"|).""("&SplitIP(1)"|).""("&SplitIP(2)"|).""("&SplitIP(3)"|)"  
     Set strMatchs=re.Execute(IP)   
      strIP=strMatchs(0).SubMatches(0) & "." & strMatchs(0).SubMatches(1)& "." & strMatchs(0).SubMatches(2)& "." & strMatchs(0).SubMatches(3)   
     if strIP=IP then MatchIP=true:exit function   
     Set strMatchs=Nothing  
     Set re=Nothing  
 next    
end function   
    
'*************************************   
'Àò±oµù¥U½X   
'*************************************     
Function getcode()    
    getcode= "<img src=""common/getcode.asp"" alt="""" style=""margin-right:40px;""/>"         
End Function  
  
'*************************************   
'­­¨î¤W¶Ç¤å¥óÃþ«¬   
'*************************************     
Function IsvalidFile(File_Type)   
    IsvalidFile = False  
    Dim GName   
    For Each GName in UP_FileType   
        If File_Type = GName Then  
            IsvalidFile = True  
            Exit For  
        End If  
    Next  
End Function  
  
'*************************************   
'ÀË´ú¬O§_¥u¥]§t­^¤å©M¼Æ¦r   
'*************************************    
Function IsValidChars(str)   
    Dim re,chkstr   
    Set re=new RegExp   
    re.IgnoreCase =true   
    re.Global=True  
    re.Pattern="[^_\.a-zA-Z\d]"  
    IsValidChars=True  
    chkstr=re.Replace(str,"")   
    if chkstr<>str then IsValidChars=False  
    set re=nothing   
End Function  
  
'*************************************   
'ÀË´ú¬O§_¥u¥]§t­^¤å©M¼Æ¦r   
'*************************************    
Function IsvalidValue(ArrayN,Str)   
    IsvalidValue = false   
    Dim GName   
    For Each GName in ArrayN   
        If Str = GName Then  
             IsvalidValue = true   
            Exit For  
        End If  
    Next  
End Function    
  
'*************************************   
'ÀË´ú¬O§_¦³®Äªº¼Æ¦r   
'*************************************   
Function IsInteger(Para)    
    IsInteger=False  
    If Not (IsNull(Para) or Trim(Para)="" or Not IsNumeric(Para)) Then  
        IsInteger=True  
    End If  
End Function  
  
'*************************************   
'¥Î¤á¦WÀË´ú   
'*************************************   
Function IsValidUserName(byVal UserName)   
    on error resume next   
    Dim i,c   
    Dim VUserName   
    IsValidUserName = True  
    For i = 1 To Len(UserName)   
        c = Lcase(Mid(UserName, i, 1))   
        If InStr("$!<>?#^%@~`&*();:+='""’Â   ", c) > 0 Then  
                IsValidUserName = False  
                Exit Function  
        End IF   
    Next  
    For Each VUserName in Register_UserName   
        If UserName = VUserName Then  
            IsValidUserName = False  
            Exit For  
        End If  
    Next  
End Function  
  
'*************************************   
'ÀË´ú¬O§_¦³®ÄªºE-mail¦a§}   
'*************************************   
Function IsValidEmail(Email)    
    Dim names, name, i, c   
    IsValidEmail = True  
    Names = Split(email, "@")   
    If UBound(names) <> 1 Then  
        IsValidEmail = False  
        Exit Function  
    End If  
    For Each name IN names   
        If Len(name) <= 0 Then  
            IsValidEmail = False  
            Exit Function  
        End If  
        For i = 1 to Len(name)   
            c = Lcase(Mid(name, i, 1))   
            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then  
                IsValidEmail = false   
                Exit Function  
            End If  
        Next  
        If Left(name, 1) = "." or Right(name, 1) = "." Then  
            IsValidEmail = false   
            Exit Function  
        End If  
    Next  
    If InStr(names(1), ".") <= 0 Then  
        IsValidEmail = False  
        Exit Function  
    End If  
    i = Len(names(1)) - InStrRev(names(1), ".")   
    If i <> 2 And i <> 3 Then  
        IsValidEmail = False  
        Exit Function  
    End If  
    If InStr(email, "..") > 0 Then  
        IsValidEmail = False  
    End If  
End Function  
  
'*************************************   
'¥[«GÃöÁä¦r   
'*************************************   
Function highlight(byVal strContent,byRef arrayWords)   
    Dim intCounter,strTemp,intPos,intTagLength,intKeyWordLength,bUpdate   
    if len(arrayWords)<1 then highlight=strContent:exit function   
    For intPos = 1 to Len(strContent)   
        bUpdate = False  
        If Mid(strContent, intPos, 1) = "<" Then  
            On Error Resume Next  
            intTagLength = (InStr(intPos, strContent, ">", 1) - intPos)   
            if err then   
              highlight=strContent   
              err.clear   
            end if   
            strTemp = strTemp & Mid(strContent, intPos, intTagLength)   
            intPos = intPos + intTagLength   
        End If  
            If arrayWords <> "" Then  
                intKeyWordLength = Len(arrayWords)   
                If LCase(Mid(strContent, intPos, intKeyWordLength)) = LCase(arrayWords) Then  
                    strTemp = strTemp & "<span class=""high1"">" & Mid(strContent, intPos, intKeyWordLength) & "</span>"  
                    intPos = intPos + intKeyWordLength - 1   
                    bUpdate = True  
                End If  
            End If  
        If bUpdate = False Then  
            strTemp = strTemp & Mid(strContent, intPos, 1)   
        End If  
    Next  
    highlight = strTemp   
End Function  
  
'*************************************   
'¹LÂo¶WÃì±µ   
'*************************************   
Function checkURL(ByVal ChkStr)   
    Dim str:str=ChkStr   
    str=Trim(str)   
    If IsNull(str) Then  
        checkURL = ""  
        Exit Function    
    End If  
    Dim re   
    Set re=new RegExp   
    re.IgnoreCase =True  
    re.Global=True  
    re.Pattern="(d)(ocument\.cookie)"  
    Str = re.replace(Str,"$1ocument cookie")   
    re.Pattern="(d)(ocument\.write)"  
    Str = re.replace(Str,"$1ocument write")   
    re.Pattern="(s)(cript:)"  
    Str = re.replace(Str,"$1cript ")   
    re.Pattern="(s)(cript)"  
    Str = re.replace(Str,"$1cript")   
    re.Pattern="(o)(bject)"  
    Str = re.replace(Str,"$1bject")   
    re.Pattern="(a)(pplet)"  
    Str = re.replace(Str,"$1pplet")   
    re.Pattern="(e)(mbed)"  
    Str = re.replace(Str,"$1mbed")   
    Set re=Nothing  
    Str = Replace(Str, ">", ">")   
    Str = Replace(Str, "<", "<")   
    checkURL=Str       
end function   
  
'*************************************   
'¹LÂo¤å¥ó¦W¦r   
'*************************************   
Function FixName(UpFileExt)   
    If IsEmpty(UpFileExt) Then Exit Function  
    FixName = Ucase(UpFileExt)   
    FixName = Replace(FixName,Chr(0),"")   
    FixName = Replace(FixName,".","")   
    FixName = Replace(FixName,"ASP","")   
    FixName = Replace(FixName,"ASA","")   
    FixName = Replace(FixName,"ASPX","")   
    FixName = Replace(FixName,"CER","")   
    FixName = Replace(FixName,"CDX","")   
    FixName = Replace(FixName,"HTR","")   
End Function  
  
'*************************************   
'¹LÂo¯S®í¦r²Å   
'*************************************   
Function CheckStr(byVal ChkStr)    
    Dim Str:Str=ChkStr   
    If IsNull(Str) Then  
        CheckStr = ""  
        Exit Function    
    End If  
    Str = Replace(Str, "&", "&")   
    Str = Replace(Str,"'","'")   
    Str = Replace(Str,"""",""")   
    Dim re   
    Set re=new RegExp   
    re.IgnoreCase =True  
    re.Global=True  
    re.Pattern="(w)(here)"  
    Str = re.replace(Str,"$1here")   
    re.Pattern="(s)(elect)"  
    Str = re.replace(Str,"$1elect")   
    re.Pattern="(i)(nsert)"  
    Str = re.replace(Str,"$1nsert")   
    re.Pattern="(c)(reate)"  
    Str = re.replace(Str,"$1reate")   
    re.Pattern="(d)(rop)"  
    Str = re.replace(Str,"$1rop")   
    re.Pattern="(a)(lter)"  
    Str = re.replace(Str,"$1lter")   
    re.Pattern="(d)(elete)"  
    Str = re.replace(Str,"$1elete")   
    re.Pattern="(u)(pdate)"  
    Str = re.replace(Str,"$1pdate")   
    re.Pattern="(\s)(or)"  
    Str = re.replace(Str,"$1or")   
    Set re=Nothing  
    CheckStr=Str   
End Function  
  
'*************************************   
'«ì´_¯S®í¦r²Å   
'*************************************   
Function UnCheckStr(ByVal Str)   
        If IsNull(Str) Then  
            UnCheckStr = ""  
            Exit Function    
        End If  
        Str = Replace(Str,"'","'")   
        Str = Replace(Str,""","""")   
        Dim re   
        Set re=new RegExp   
        re.IgnoreCase =True  
        re.Global=True  
        re.Pattern="(w)(here)"  
        str = re.replace(str,"$1here")   
        re.Pattern="(s)(elect)"  
        str = re.replace(str,"$1elect")   
        re.Pattern="(i)(nsert)"  
        str = re.replace(str,"$1nsert")   
        re.Pattern="(c)(reate)"  
        str = re.replace(str,"$1reate")   
        re.Pattern="(d)(rop)"  
        str = re.replace(str,"$1rop")   
        re.Pattern="(a)(lter)"  
        str = re.replace(str,"$1lter")   
        re.Pattern="(d)(elete)"  
        str = re.replace(str,"$1elete")   
        re.Pattern="(u)(pdate)"  
        str = re.replace(str,"$1pdate")   
        re.Pattern="(\s)(or)"  
        Str = re.replace(Str,"$1or")   
        Set re=Nothing  
        Str = Replace(Str, "&", "&")   
        UnCheckStr=Str   
End Function  
  
'*************************************   
'Âà´«HTML¥N½X   
'*************************************   
Function HTMLEncode(ByVal reString)    
    Dim Str:Str=reString   
    If Not IsNull(Str) Then  
        Str = Replace(Str, ">", ">")   
        Str = Replace(Str, "<", "<")   
        Str = Replace(Str, CHR(9), "    ")   
        Str = Replace(Str, CHR(32), " ")   
        Str = Replace(Str, CHR(39), "'")   
        Str = Replace(Str, CHR(34), """)   
        Str = Replace(Str, CHR(13), "")   
        Str = Replace(Str, CHR(10), "<br/>")   
        HTMLEncode = Str   
    End If  
End Function  
  
'*************************************   
'¤ÏÂà´«HTML¥N½X   
'*************************************   
Function HTMLDecode(ByVal reString)    
    Dim Str:Str=reString   
    If Not IsNull(Str) Then  
        Str = Replace(Str, ">", ">")   
        Str = Replace(Str, "<", "<")   
        Str = Replace(Str, "    ", CHR(9))   
        Str = Replace(Str, " ", CHR(32))   
        Str = Replace(Str, "'", CHR(39))   
        Str = Replace(Str, """, CHR(34))   
        Str = Replace(Str, "", CHR(13))   
        Str = Replace(Str, "<br/>", CHR(10))   
        HTMLDecode = Str   
    End If  
End Function  
  
'*************************************   
'«ì´_&¦r²Å   
'*************************************   
function ClearHTML(ByVal reString)   
    Dim Str:Str=reString   
    If Not IsNull(Str) Then  
        Str = Replace(Str, "&", "&")   
        ClearHTML = Str   
    End If  
End Function  
  
'*************************************   
'¹LÂotextarea   
'*************************************   
Function UBBFilter(ByVal reString)   
    Dim Str:Str=reString   
    If Not IsNull(Str) Then  
        Str = Replace(Str, "</textarea>", "</textarea>")   
        UBBFilter = Str   
    End If  
End Function  
  
'*************************************   
'¹LÂoHTML¥N½X   
'*************************************   
Function EditDeHTML(byVal Content)   
    EditDeHTML=Content   
    IF Not IsNull(EditDeHTML) Then  
        EditDeHTML=UnCheckStr(EditDeHTML)   
        EditDeHTML=Replace(EditDeHTML,"&","&")   
        EditDeHTML=Replace(EditDeHTML,"<","<")   
        EditDeHTML=Replace(EditDeHTML,">",">")   
        EditDeHTML=Replace(EditDeHTML,chr(34),""")   
        EditDeHTML=Replace(EditDeHTML,chr(39),"'")   
    End IF   
End Function  
  
'*************************************   
'¤é´ÁÂà´«¨ç¼Æ   
'*************************************   
Function DateToStr(DateTime,ShowType)     
    Dim DateMonth,DateDay,DateHour,DateMinute,DateWeek,DateSecond   
    Dim FullWeekday,shortWeekday,Fullmonth,Shortmonth,TimeZone1,TimeZone2   
    TimeZone1="+0800"  
    TimeZone2="+08:00"  
    FullWeekday=Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")   
    shortWeekday=Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")   
    Fullmonth=Array("January","February","March","April","May","June","July","August","September","October","November","December")   
    Shortmonth=Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")   
  
    DateMonth=Month(DateTime)   
    DateDay=Day(DateTime)   
    DateHour=Hour(DateTime)   
    DateMinute=Minute(DateTime)   
    DateWeek=weekday(DateTime)   
    DateSecond=Second(DateTime)   
    If Len(DateMonth)<2 Then DateMonth="0"&DateMonth   
    If Len(DateDay)<2 Then DateDay="0"&DateDay   
    If Len(DateMinute)<2 Then DateMinute="0"&DateMinute   
    Select Case ShowType   
    Case "Y-m-d"     
        DateToStr=Year(DateTime)"-"&DateMonth"-"&DateDay   
    Case "Y-m-d H:I A"  
        Dim DateAMPM   
        If DateHour>12 Then    
            DateHour=DateHour-12   
            DateAMPM="PM"  
        Else  
            DateHour=DateHour   
            DateAMPM="AM"  
        End If  
        If Len(DateHour)<2 Then DateHour="0"&DateHour      
        DateToStr=Year(DateTime)"-"&DateMonth"-"&DateDay" "&DateHour":"&DateMinute" "&DateAMPM   
    Case "Y-m-d H:I:S"  
        If Len(DateHour)<2 Then DateHour="0"&DateHour      
        If Len(DateSecond)<2 Then DateSecond="0"&DateSecond   
        DateToStr=Year(DateTime)"-"&DateMonth"-"&DateDay" "&DateHour":"&DateMinute":"&DateSecond   
    Case "YmdHIS"  
        DateSecond=Second(DateTime)   
        If Len(DateHour)<2 Then DateHour="0"&DateHour      
        If Len(DateSecond)<2 Then DateSecond="0"&DateSecond   
        DateToStr=Year(DateTime)&DateMonth&DateDay&DateHour&DateMinute&DateSecond      
    Case "ym"  
        DateToStr=Right(Year(DateTime),2)&DateMonth   
    Case "d"  
        DateToStr=DateDay   
    Case "ymd"  
        DateToStr=Right(Year(DateTime),4)&DateMonth&DateDay   
    Case "mdy"    
        Dim DayEnd   
        select Case DateDay   
         Case 1    
          DayEnd="st"  
         Case 2   
          DayEnd="nd"  
         Case 3   
          DayEnd="rd"  
         Case Else  
          DayEnd="th"  
        End Select    
        DateToStr=Fullmonth(DateMonth-1)" "&DateDay&DayEnd" "&Right(Year(DateTime),4)   
    Case "w,d m y H:I:S"    
        DateSecond=Second(DateTime)   
        If Len(DateHour)<2 Then DateHour="0"&DateHour      
        If Len(DateSecond)<2 Then DateSecond="0"&DateSecond   
        DateToStr=shortWeekday(DateWeek-1)","&DateDay" "& Left(Fullmonth(DateMonth-1),3) " "&Right(Year(DateTime),4)" "&DateHour":"&DateMinute":"&DateSecond" "&TimeZone1   
    Case "y-m-dTH:I:S"  
        If Len(DateHour)<2 Then DateHour="0"&DateHour      
        If Len(DateSecond)<2 Then DateSecond="0"&DateSecond   
        DateToStr=Year(DateTime)"-"&DateMonth"-"&DateDay"T"&DateHour":"&DateMinute":"&DateSecond&TimeZone2   
    Case Else  
        If Len(DateHour)<2 Then DateHour="0"&DateHour   
        DateToStr=Year(DateTime)"-"&DateMonth"-"&DateDay" "&DateHour":"&DateMinute   
    End Select  
End Function  
  
'*************************************   
'¤À­¶¨ç¼Æ   
'*************************************   
dim FirstShortCut,ShortCut   
FirstShortCut=false   
Function MultiPage(Numbers,Perpage,Curpage,Url_Add,aname,Style)    
    CurPage=Int(Curpage)   
    Numbers=Int(Numbers)   
    Dim URL   
    URL=Request.ServerVariables("Script_Name")&Url_Add   
    MultiPage=""  
    Dim Page,Offset,PageI   
'   If Int(Numbers)>Int(PerPage) Then   
        Page=9   
        Offset=4   
        Dim Pages,FromPage,ToPage   
        If Numbers Mod Cint(Perpage)=0 Then  
            Pages=Int(Numbers/Perpage)   
        Else  
            Pages=Int(Numbers/Perpage)+1   
        End If  
        FromPage=Curpage-Offset   
        ToPage=Curpage+Page-Offset-1   
        If Page>Pages Then  
            FromPage=1   
            ToPage=Pages   
        Else  
            If FromPage<1 Then  
                Topage=Curpage+1-FromPage   
                FromPage=1   
                If (ToPage-FromPage)<Page And (ToPage-FromPage)<Pages Then ToPage=Page   
            ElseIF Topage>Pages Then  
                FromPage =Curpage-Pages +ToPage   
                ToPage=Pages   
                If (ToPage-FromPage)<Page And (ToPage-FromPage)<Pages Then FromPage=Pages-Page+1   
            End If  
        End If  
         MultiPage="<div class=""page"" style="""&Style"""><ul>"  
       'if Curpage<>1 then MultiPage=MultiPage&"<li class=""PageL""><a href="""&Url&"page=1"" class=""PageLbutton"" title=""²Ä¤@­¶""></a></li>"   
        MultiPage=MultiPage"<li class=""pageNumber"">"  
        if Curpage<>1 then MultiPage=MultiPage"<a href="""&Url"page=1"" title=""²Ä¤@­¶"" style=""text-decoration:none""><</a> | "  
        if not FirstShortCut then ShortCut=" accesskey="",""" else ShortCut=""  
        if Curpage<>1 then MultiPage=MultiPage"<a href="""&Url"page="&CurPage-1""" title=""¤W¤@­¶"" style=""text-decoration:none;"""&ShortCut"></a>"  
        For PageI=FromPage TO ToPage   
            If PageI<>CurPage Then  
                MultiPage=MultiPage"<a href="""&Url"page="&PageI&aname""">"&PageI"</a> | "  
            Else  
                MultiPage=MultiPage"<strong>"&PageI"</strong>"  
                if PageI<>Pages then MultiPage=MultiPage" | "  
            End If  
        Next  
        if not FirstShortCut then ShortCut=" accesskey="".""" else ShortCut=""  
        if Curpage<>pages then MultiPage=MultiPage"<a href="""&Url"page="&CurPage+1""" title=""¤U¤@­¶"" style=""text-decoration:none"""&ShortCut"></a>"  
        if Curpage<>pages then MultiPage=MultiPage"<a href="""&Url"page="&Pages&aname""" title=""³Ì«á¤@­¶"" style=""text-decoration:none"">></a>"  
        MultiPage=MultiPage"</li>"  
        'If Int(Pages)>Int(Page) Then   
        '   MultiPage=MultiPage&"<li>...</li><li><a href="""&Url&"page="&Pages&aname&""">"&pages&"</a></li>"   
        'End If   
        'if Curpage<>pages then MultiPage=MultiPage&"<li class=""PageR""><a href="""&Url&"page="&Pages&aname&""" class=""PageRbutton"" title=""³Ì«á¤@­¶""></a></li>"   
        MultiPage=MultiPage"</ul></div>"  
'   End If   
FirstShortCut=true   
End Function  
  
'*************************************   
'¤Á³Î¤º®e - «ö¦æ¤À³Î   
'*************************************   
Function SplitLines(byVal Content,byVal ContentNums)    
    Dim ts,i,l   
    ContentNums=int(ContentNums)   
    If IsNull(Content) Then Exit Function  
    i=1   
    ts = 0   
    For i=1 to Len(Content)   
      l=Lcase(Mid(Content,i,5))   
        If l="<br/>" Then  
            ts=ts+1   
        End If  
      l=Lcase(Mid(Content,i,4))   
        If l="<br>" Then  
            ts=ts+1   
        End If  
      l=Lcase(Mid(Content,i,3))   
        If l="<p>" Then  
            ts=ts+1   
        End If  
    If ts>ContentNums Then Exit For    
    Next  
    If ts>ContentNums Then  
        Content=Left(Content,i-1)   
    End If  
    SplitLines=Content   
End Function  
  
'*************************************   
'¤Á³Î¤º®e - «ö¦r²Å¤À³Î   
'*************************************   
Function CutStr(byVal Str,byVal StrLen)   
    Dim l,t,c,i   
    If IsNull(Str) Then CutStr="":Exit Function  
    l=Len(str)   
    StrLen=int(StrLen)   
    t=0   
    For i=1 To l   
        c=Asc(Mid(str,i,1))   
        If c<0 or c>255 Then t=t+2 Else t=t+1   
        IF t>=StrLen Then  
            CutStr=left(Str,i)"..."  
            Exit For  
        Else  
            CutStr=Str   
        End If  
    Next  
End Function  
  
'*************************************   
'§R°£¤Þ¥Î¼ÐÅÒ   
'*************************************   
Function DelQuote(strContent)   
    If IsNull(strContent) Then Exit Function  
    Dim re   
    Set re=new RegExp   
    re.IgnoreCase =True  
    re.Global=True  
    re.Pattern="\[quote\](.[^\]]*?)\[\/quote\]"  
    strContent= re.Replace(strContent,"")   
    re.Pattern="\[quote=(.[^\]]*)\](.[^\]]*?)\[\/quote\]"  
    strContent= re.Replace(strContent,"")   
    Set re=Nothing  
    DelQuote=strContent   
End Function  
  
'*************************************   
'Àò¨ú«È¤áºÝIP   
'*************************************   
function getIP()    
         dim strIP,IP_Ary,strIP_list   
         strIP_list=Replace(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),"'","")   
            
         If InStr(strIP_list,",")<>0 Then  
            IP_Ary = Split(strIP_list,",")   
            strIP = IP_Ary(0)   
         Else  
            strIP = strIP_list   
         End IF   
            
         If strIP=Empty Then strIP=Replace(Request.ServerVariables("REMOTE_ADDR"),"'","")   
         getIP=strIP   
End Function  
  
'*************************************   
'Àò¨ú«È¤áºÝÂsÄý¾¹«H®§   
'*************************************   
function getBrowser(strUA)    
 dim arrInfo,strType,temp1,temp2   
 strType=""  
 strUA=LCase(strUA)   
 arrInfo=Array("Unkown","Unkown")   
 'ÂsÄý¾¹§PÂ_   
    if Instr(strUA,"mozilla")>0 then arrInfo(0)="Mozilla"  
    if Instr(strUA,"icab")>0 then arrInfo(0)="iCab"  
    if Instr(strUA,"lynx")>0 then arrInfo(0)="Lynx"  
    if Instr(strUA,"links")>0 then arrInfo(0)="Links"  
    if Instr(strUA,"elinks")>0 then arrInfo(0)="ELinks"  
    if Instr(strUA,"jbrowser")>0 then arrInfo(0)="JBrowser"  
    if Instr(strUA,"konqueror")>0 then arrInfo(0)="konqueror"  
    if Instr(strUA,"wget")>0 then arrInfo(0)="wget"  
    if Instr(strUA,"ask jeeves")>0 or Instr(strUA,"teoma")>0 then arrInfo(0)="Ask Jeeves/Teoma"  
    if Instr(strUA,"wget")>0 then arrInfo(0)="wget"  
    if Instr(strUA,"opera")>0 then arrInfo(0)="opera"  
  
    if Instr(strUA,"gecko")>0 then    
      strType="[Gecko]"  
      arrInfo(0)="Mozilla"  
      if Instr(strUA,"aol")>0 then arrInfo(0)="AOL"  
      if Instr(strUA,"netscape")>0 then arrInfo(0)="Netscape"  
      if Instr(strUA,"firefox")>0 then arrInfo(0)="FireFox"  
      if Instr(strUA,"chimera")>0 then arrInfo(0)="Chimera"  
      if Instr(strUA,"camino")>0 then arrInfo(0)="Camino"  
      if Instr(strUA,"galeon")>0 then arrInfo(0)="Galeon"  
      if Instr(strUA,"k-meleon")>0 then arrInfo(0)="K-Meleon"  
      arrInfo(0)=arrInfo(0)+strType   
   end if   
      
   if Instr(strUA,"bot")>0 or Instr(strUA,"crawl")>0 then    
      strType="[Bot/Crawler]"  
      arrInfo(0)=""  
      if Instr(strUA,"grub")>0 then arrInfo(0)="Grub"  
      if Instr(strUA,"googlebot")>0 then arrInfo(0)="GoogleBot"  
      if Instr(strUA,"msnbot")>0 then arrInfo(0)="MSN Bot"  
      if Instr(strUA,"slurp")>0 then arrInfo(0)="Yahoo! Slurp"  
      arrInfo(0)=arrInfo(0)+strType   
  end if   
     
  if Instr(strUA,"applewebkit")>0 then    
      strType="[AppleWebKit]"  
      arrInfo(0)=""  
      if Instr(strUA,"omniweb")>0 then arrInfo(0)="OmniWeb"  
      if Instr(strUA,"safari")>0 then arrInfo(0)="Safari"  
      arrInfo(0)=arrInfo(0)+strType   
  end if    
     
  if Instr(strUA,"msie")>0 then    
      strType="[MSIE"  
      temp1=mid(strUA,(Instr(strUA,"msie")+4),6)   
      temp2=Instr(temp1,";")   
      temp1=left(temp1,temp2-1)   
      strType=strType & temp1 "]"  
      arrInfo(0)="Internet Explorer"  
      if Instr(strUA,"msn")>0 then arrInfo(0)="MSN"  
      if Instr(strUA,"aol")>0 then arrInfo(0)="AOL"  
      if Instr(strUA,"webtv")>0 then arrInfo(0)="WebTV"  
      if Instr(strUA,"myie2")>0 then arrInfo(0)="MyIE2"  
      if Instr(strUA,"maxthon")>0 then arrInfo(0)="Maxthon"  
      if Instr(strUA,"gosurf")>0 then arrInfo(0)="GoSurf"  
      if Instr(strUA,"netcaptor")>0 then arrInfo(0)="NetCaptor"  
      if Instr(strUA,"sleipnir")>0 then arrInfo(0)="Sleipnir"  
      if Instr(strUA,"avant browser")>0 then arrInfo(0)="AvantBrowser"  
      if Instr(strUA,"greenbrowser")>0 then arrInfo(0)="GreenBrowser"  
      if Instr(strUA,"slimbrowser")>0 then arrInfo(0)="SlimBrowser"  
      arrInfo(0)=arrInfo(0)+strType   
   end if   
    
 '¾Þ§@¨t²Î§PÂ_   
    if Instr(strUA,"windows")>0 then arrInfo(1)="Windows"  
    if Instr(strUA,"windows ce")>0 then arrInfo(1)="Windows CE"  
    if Instr(strUA,"windows 95")>0 then arrInfo(1)="Windows 95"  
    if Instr(strUA,"win98")>0 then arrInfo(1)="Windows 98"  
    if Instr(strUA,"windows 98")>0 then arrInfo(1)="Windows 98"  
    if Instr(strUA,"windows 2000")>0 then arrInfo(1)="Windows 2000"  
    if Instr(strUA,"windows xp")>0 then arrInfo(1)="Windows XP"  
  
    if Instr(strUA,"windows nt")>0 then   
      arrInfo(1)="Windows NT"  
      if Instr(strUA,"windows nt 5.0")>0 then arrInfo(1)="Windows 2000"  
      if Instr(strUA,"windows nt 5.1")>0 then arrInfo(1)="Windows XP"  
      if Instr(strUA,"windows nt 5.2")>0 then arrInfo(1)="Windows 2003"  
    end if   
    if Instr(strUA,"x11")>0 or Instr(strUA,"unix")>0 then arrInfo(1)="Unix"  
    if Instr(strUA,"sunos")>0 or Instr(strUA,"sun os")>0 then arrInfo(1)="SUN OS"  
    if Instr(strUA,"powerpc")>0 or Instr(strUA,"ppc")>0 then arrInfo(1)="PowerPC"  
    if Instr(strUA,"macintosh")>0 then arrInfo(1)="Mac"  
    if Instr(strUA,"mac osx")>0 then arrInfo(1)="MacOSX"  
    if Instr(strUA,"freebsd")>0 then arrInfo(1)="FreeBSD"  
    if Instr(strUA,"linux")>0 then arrInfo(1)="Linux"  
    if Instr(strUA,"palmsource")>0 or Instr(strUA,"palmos")>0 then arrInfo(1)="PalmOS"  
    if Instr(strUA,"wap ")>0 then arrInfo(1)="WAP"  
     
 'arrInfo(0)=strUA    
 getBrowser=arrInfo   
end function   
  
'*************************************   
'­pºâÀH¾÷¼Æ   
'*************************************   
function randomStr(intLength)   
    dim strSeed,seedLength,pos,str,i   
    strSeed = "abcdefghijklmnopqrstuvwxyz1234567890"  
    seedLength=len(strSeed)   
    str=""  
    Randomize   
    for i=1 to intLength   
     str=str+mid(strSeed,int(seedLength*rnd)+1,1)   
    next   
    randomStr=str   
end function   
  
'*************************************   
'¦Û°Ê³¬¦XUBB   
'*************************************   
function closeUBB(strContent)   
  dim arrTags,i,OpenPos,ClosePos,re,strMatchs,j,Match   
    Set re=new RegExp   
    re.IgnoreCase =True  
    re.Global=True  
    arrTags=array("code","quote","list","color","align","font","size","b","i","u","html")   
  for i=0 to ubound(arrTags)   
   OpenPos=0   
   ClosePos=0   
      
   re.Pattern="\["+arrTags(i)+"(=[^\[\]]+|)\]"  
   Set strMatchs=re.Execute(strContent)   
   For Each Match in strMatchs   
    OpenPos=OpenPos+1   
   next   
   re.Pattern="\[/"+arrTags(i)+"\]"  
   Set strMatchs=re.Execute(strContent)   
   For Each Match in strMatchs   
    ClosePos=ClosePos+1   
   next   
   for j=1 to OpenPos-ClosePos   
      strContent=strContent+"[/"+arrTags(i)+"]"  
   next   
  next   
closeUBB=strContent   
end function   
  
'*************************************   
'¦Û°Ê³¬¦XHTML   
'*************************************   
function closeHTML(strContent)   
  dim arrTags,i,OpenPos,ClosePos,re,strMatchs,j,Match   
    Set re=new RegExp   
    re.IgnoreCase =True  
    re.Global=True  
    arrTags=array("p","div","span","table","ul","font","b","u","i","h1","h2","h3","h4","h5","h6")   
  for i=0 to ubound(arrTags)   
   OpenPos=0   
   ClosePos=0   
      
   re.Pattern="\<"+arrTags(i)+"( [^\<\>]+|)\>"  
   Set strMatchs=re.Execute(strContent)   
   For Each Match in strMatchs   
    OpenPos=OpenPos+1   
   next   
   re.Pattern="\</"+arrTags(i)+"\>"  
   Set strMatchs=re.Execute(strContent)   
   For Each Match in strMatchs   
    ClosePos=ClosePos+1   
   next   
   for j=1 to OpenPos-ClosePos   
      strContent=strContent+"</"+arrTags(i)+">"  
   next   
  next   
closeHTML=strContent   
end function   
  
'*************************************   
'Åª¨ú¤å¥ó   
'*************************************   
Function LoadFromFile(ByVal File)   
    Dim objStream   
    Dim RText   
    RText=array(0,"")   
    On Error Resume Next  
    Set objStream = Server.CreateObject("ADODB.Stream")   
    If Err Then    
        RText=array(Err.Number,Err.Description)   
        LoadFromFile=RText   
        Err.Clear   
        exit function   
    End If  
    With objStream   
        .Type = 2   
        .Mode = 3   
        .Open   
        .Charset = "utf-8"  
        .Position = objStream.Size   
        .LoadFromFile Server.MapPath(File)   
        If Err.Number<>0 Then  
           RText=array(Err.Number,Err.Description)   
           LoadFromFile=RText   
           Err.Clear   
           exit function   
        End If  
        RText=array(0,.ReadText)   
        .Close   
    End With  
    LoadFromFile=RText   
    Set objStream = Nothing  
End Function  
  
'*************************************   
'«O¦s¤å¥ó   
'*************************************   
Function SaveToFile(ByVal strBody,ByVal File)   
    Dim objStream   
    Dim RText   
    RText=array(0,"")   
    On Error Resume Next  
    Set objStream = Server.CreateObject("ADODB.Stream")   
    If Err Then    
        RText=array(Err.Number,Err.Description)   
        Err.Clear   
        exit function   
    End If  
    With objStream   
        .Type = 2   
        .Open   
        .Charset = "utf-8"  
        .Position = objStream.Size   
        .WriteText = strBody   
        .SaveToFile Server.MapPath(File),2   
        .Close   
    End With  
    RText=array(0,"«O¦s¤å¥ó¦¨¥\!")   
    SaveToFile=RText   
    Set objStream = Nothing  
End Function  
  
'*************************************   
'¼Æ¾Ú®w²K¥[­×§ï¾Þ§@   
'*************************************   
function DBQuest(table,DBArray,Action)   
 dim AddCount,TempDB,i,v   
 if Action<>"insert" or Action<>"update" then Action="insert"  
 if Action="insert" then v=2 else v=3   
 if not IsArray(DBArray) then   
   DBQuest=-1   
   exit function   
 else   
   Set TempDB=Server.CreateObject("ADODB.RecordSet")   
   On Error Resume Next  
   TempDB.Open table,Conn,1,v   
   if err then   
    DBQuest=-2   
    exit function   
   end if   
   if Action="insert" then TempDB.addNew   
   AddCount=UBound(DBArray,1)   
   for i=0 to AddCount   
    TempDB(DBArray(i)(0))=DBArray(i)(1)   
   next   
   TempDB.update   
   TempDB.close   
   set TempDB=nothing   
   DBQuest=0   
 end if   
end Function  
  
'*************************************   
'ÀË´ú¨t²Î²Õ¥ó¬O§_¦w¸Ë   
'*************************************   
Function CheckObjInstalled(strClassString)   
    On Error Resume Next  
    Dim Temp   
    Err = 0   
    Dim TmpObj   
    Set TmpObj = Server.CreateObject(strClassString)   
    Temp = Err   
    IF Temp = 0 or Temp = -2147221477 Then  
        CheckObjInstalled=true   
    ElseIF Temp = 1 or Temp = -2147221005 Then  
        CheckObjInstalled=false   
    End IF   
    Err.Clear   
    Set TmpObj = Nothing  
    Err = 0   
End Function  
  
'*************************************   
'§PÂ_ªA°È¾¹Microsoft.XMLDOM   
'*************************************   
Function getXMLDOM   
    On Error Resume Next  
    Dim Temp   
    getXMLDOM="Microsoft.XMLDOM"  
    Err = 0   
    Dim TmpObj   
    Set TmpObj = Server.CreateObject(getXMLDOM)   
    Temp = Err   
    IF Temp = 1 or Temp = -2147221005 Then  
        getXMLDOM="Msxml2.DOMDocument.5.0"  
    End IF   
    Err.Clear   
    Set TmpObj = Nothing  
    Err = 0   
end function   
  
'*************************************   
'§PÂ_ªA°È¾¹MSXML2.ServerXMLHTTP   
'*************************************   
Function getXMLHTTP   
    On Error Resume Next  
    Dim Temp   
    getXMLHTTP="MSXML2.ServerXMLHTTP"  
    Err = 0   
    Dim TmpObj   
    Set TmpObj = Server.CreateObject(getXMLHTTP)   
    Temp = Err   
    IF Temp = 1 or Temp = -2147221005 Then  
        getXMLHTTP="Msxml2.ServerXMLHTTP.5.0"  
    End IF   
    Err.Clear   
    Set TmpObj = Nothing  
    Err = 0   
end function   
  
'*************************************   
'©U§£ÃöÁä¦r¹LÂo   
'*************************************   
function filterSpam(str,path)   
  on error resume next   
     filterSpam = false   
     dim spamXml,spamItem   
     Set spamXml = Server.CreateObject(getXMLDOM)   
       If Err Then     
           Err.clear   
           exit function   
       end if   
     spamXml.async = false     
     spamXml.load(Server.MapPath(path))   
     if spamXml.parseerror.errorcode=0 then   
       For Each spamItem in spamXml.selectNodes("//key")   
            if InStr(Lcase(str),Lcase(spamItem.text))<>0 then   
               filterSpam = true   
               exit function   
            end if   
       next   
     end if   
     set spamXml=nothing   
end function   
  
'*********************************************************   
' ¥Øªº¡G    ÀË¬d¥¿«h¦¡   
' ¿é¤J¡G    id   
' ªð¦^¡G    ¦¨¥\¬°True   
'*********************************************************   
Function CheckRegExp(source,para)   
  
    If para="[username]" Then  
        para="^[.A-Za-z0-9\u4e00-\u9fa5]+$"  
    End If  
    If para="[password]" Then  
        para="^[a-z0-9]+$"  
    End If  
    If para="[email]" Then  
        para="^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*\.)+[a-zA-Z]*)$"  
    End If  
    If para="[homepage]" Then  
        para="^[a-zA-Z]+://[a-zA-z0-9\-\./]+?/*$"  
    End If  
    If para="[nojapan]" Then  
        para="[\u3040-\u30ff]+"  
    End If  
    If para="[guid]" Then  
        para="^\w{8}\-\w{4}\-\w{4}\-\w{4}\-\w{12}$"  
    End If  
  
    Dim re   
    Set re = New RegExp   
    re.Global = True  
    re.Pattern = para   
    re.IgnoreCase = False  
    CheckRegExp = re.Test(source)   
  
End Function  
  
'**********************************************   
'Àò¨ú¦b½u¤H¼Æ   
'**********************************************   
function getOnline   
    getOnline=1   
    if len(Application(space_CookieName"_onlineCount"))>0 then   
        if DateDiff("s",Application(space_CookieName"_userOnlineCountTime"),now())>60 then   
                Application.Lock()   
                Application(space_CookieName"_online")=Application(space_CookieName"_onlineCount")   
                Application(space_CookieName"_onlineCount")=1   
                Application(space_CookieName"_onlineCountKey")=randStr(2)   
                Application(space_CookieName"_userOnlineCountTime")=now()   
                Application.Unlock()   
        else   
                if Session(space_CookieName"userOnlineKey")<>Application(space_CookieName"_onlineCountKey") then   
                    Application.Lock()   
                    Application(space_CookieName"_onlineCount")=Application(space_CookieName"_onlineCount")+1   
                    Application.Unlock()   
                    Session(space_CookieName"userOnlineKey")=Application(space_CookieName"_onlineCountKey")   
                end if   
        end if   
    else   
        Application.Lock   
        Application(space_CookieName"_online")=1   
        Application(space_CookieName"_onlineCount")=1   
        Application(space_CookieName"_onlineCountKey")=randStr(2)   
        Application(space_CookieName"_userOnlineCountTime")=now()   
        Application.Unlock   
    end if   
    getOnline=Application(space_CookieName"_online")   
end Function  
  
%>  