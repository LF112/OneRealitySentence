dim word, Filer

Filer="Rua.txt"

word=ReadFile(Filer, "utf-8")
for i=1 to len(word)
data=mid(word,i,1)
Execute "x"&i&"=data"
next
for i=1 to len(word)
Call WriteToFile(i & ".txt", eval("x"&i), "utf-8")
next

Function ReadFile(FileUrl, CharSet) 
    Dim Str 
    Set stm = CreateObject("Adodb.Stream") 
    stm.Type = 2 
    stm.mode = 3 
    stm.charset = CharSet 
    stm.Open 
    stm.loadfromfile FileUrl 
    Str = stm.readtext 
    stm.Close 
    Set stm = Nothing 
    ReadFile = Str 
End Function

Function WriteToFile (FileUrl, Str, CharSet) 
    Set stm = CreateObject("Adodb.Stream") 
    stm.Type = 2 
    stm.mode = 3 
    stm.charset = CharSet 
    stm.Open 
    stm.WriteText Str 
    stm.SaveToFile FileUrl, 2 
    stm.flush 
    stm.Close 
    Set stm = Nothing 
End Function