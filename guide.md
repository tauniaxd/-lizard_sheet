[](#蜥蜴表單🦎📋使用說明-httpsbitlylizard-sheet-how "蜥蜴表單🦎📋使用說明-httpsbitlylizard-sheet-how")蜥蜴表單🦎📋使用說明 [https://bit.ly/lizard-sheet-how](https://bit.ly/lizard-sheet-how)
=====================================================================================================================================================================

> “Insanity is doing the same thing over and over again, but expecting different results.” - Albert Einstein

### [](#你想想，你在VGHTPE，打著病歷，吃著早餐，還唱著歌，突然主治醫師來了問病人Data。你還在慌亂中點開各種Data嗎？還在浪費時間搜集以下資訊嗎： "你想想，你在VGHTPE，打著病歷，吃著早餐，還唱著歌，突然主治醫師來了問病人Data。你還在慌亂中點開各種Data嗎？還在浪費時間搜集以下資訊嗎：")你想想，你在VGHTPE，打著病歷，吃著早餐，還唱著歌，突然主治醫師來了問病人Data。你還在慌亂中點開各種Data嗎？還在浪費時間搜集以下資訊嗎：

Name Sex Age  
WBC/Hb/Plt B\\N\\L INR/Aptt  
Fib\\Ddimer BUN/Cr/Na/K  
fCa\\Ca\\Mg\\iP Alb/T./D.Bilirubin  
ALT\\AST\\AlkP\\GT\\Amy\\Lip CRP/Lac/Pct  
CK\\MB\\TnI (FiO2)/pH/pO2/pCO2/HCO3  
Vitals Signs O2Device SpO2  
Weight Input/Output

### [](#你需要的是一張一鍵搜集好所有病人資訊的蜥蜴表單 "你需要的是一張一鍵搜集好所有病人資訊的蜥蜴表單")你需要的是一張一鍵搜集好所有病人資訊的蜥蜴表單

> 有任何問題及開發上的建議，可以直接在這裡留言，或直接[加我Fb好友](https://www.facebook.com/Tim.H.Lin)。若三言兩語講不完，也可來信ppoiu87@gmail.com

[](#A-環境搭建-aka-如何安裝 "A-環境搭建-aka-如何安裝")A. 環境搭建 a.k.a. 如何安裝
---------------------------------------------------------

> 有點麻煩QQ，但安裝一次永久使用，CP值極高

### [](#1-安裝SeleniumBasic---讓Excel可以用VBA跑Selenium的套件 "1-安裝SeleniumBasic---讓Excel可以用VBA跑Selenium的套件")1\. 安裝SeleniumBasic - 讓Excel可以用VBA跑Selenium的套件

> 本方法僅支援Windows

*   [下載 SeleniumBasic-2.0.9.0.exe 並安裝](https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0) [(備用載點)](https://drive.google.com/file/d/1CRSAFCKRvDTB-QM7cH0zEZy8yO7G6mcO/view?usp=sharing)
*   在`Select Components` 這一步，在選單中選擇`Compact Installation`，不要選`Full installation` (那些預設的Web driver都是過時的舊版，勿選，我們之後要安裝新版)，我們會把這個軟體安裝在這個路徑

    Install folder / Current User :
    C:\Users\電腦使用者名稱\AppData\Local\SeleniumBasic
    

### [](#2-安裝Edge-Web-Driver，讓Selenium可以操作Edge去爬網站 "2-安裝Edge-Web-Driver，讓Selenium可以操作Edge去爬網站")2\. 安裝Edge Web Driver，讓Selenium可以操作Edge去爬網站

*   將下面這行貼到搜尋列上送出查詢本台電腦Edge版本號  
    `edge://settings/help`
    
    *   如果版本號是64位元，即是x64
    *   32位元則是x86
*   [下載Edge Web Driver](https://developer.microsoft.com/zh-tw/microsoft-edge/tools/webdriver/)
    
    *   選擇對應的版本x64 or x86  
        
        ![](https://i.imgur.com/XjoGx2i.png)
*   將下載的檔案解壓縮到上面那個SeleniumBasic的安裝資料夾裡，可以直接貼上剛剛存下來的`路徑`，如果你不會用貼上路徑的方式，也可以直接Copy Paste解壓出來的檔案到剛剛那個資料夾裡
    
*   將檔名`msedgedriver`前兩個ms刪掉，改成`edgedriver`，這樣SeleniumBasic 才認得出來。  
    
    ![](https://i.imgur.com/PJSDOFv.png)
    

[](#B-表單說明 "B-表單說明")B. 表單說明
---------------------------

### [](#1-下載 "1-下載")1\. 下載

*   [表單下載連結](https://drive.google.com/drive/folders/1kJAGFNZQJuTwpNBf8GNVTJvDRJWZibtD)
*   或使用短網址：  
    `bit.ly/lizard-sheet`

### [](#2-首次開啟時，選擇啟用編輯、跟啟用巨集 "2-首次開啟時，選擇啟用編輯、跟啟用巨集")2\. 首次開啟時，選擇`啟用編輯`、跟`啟用巨集`

![](https://i.imgur.com/c6b6Udu.png)

*   這個表單開啟時，會自動監測是否有安裝Selenium，如果安裝成功，就不會跳出沒裝的警告
*   在【由此開始】表單按【登入】，會彈出系統登入畫面。登入後縮小視窗，勿關，關了就要重登
*   按【建立連線】這一步一定要做！，不然會出錯(建立Session)

### [](#3-我的病人 "3-我的病人")3\. 我的病人

*   分別按載入我的病人、載入數據，你的表單就完成了

### [](#4-病房專區 "4-病房專區")4\. 病房專區

*   在下方有病房區，打入你的病房名稱，這裡是撈選單裡的資料，例如141病房是要用A141，才會正確
*   科別一定要選
*   在病房專區裡也有小按鈕可以按，功能一樣  
    
    ![](https://i.imgur.com/qiVnUUB.png)

### [](#5-重要：沒用時記得登出並釋放記憶體 "5-重要：沒用時記得登出並釋放記憶體")5\. 重要：沒用時記得登出並釋放記憶體

這個Excel程式本質是操作Edge來撈網址，所以會等於多開一個分頁，當電腦卡時，直接把那個彈出來的Edge Driver關掉就好，就會直接登出

### [](#6-請選橫式列印 "6-請選橫式列印")6\. 請選橫式列印

*   選取你要的範圍，或按全選按鈕
*   在內容選擇的狀態下，在列印裡選擇所選範圍，這樣就會把所要的內容印出來了

[](#C-設計理念 "C-設計理念")C. 設計理念
---------------------------

*   SMAC跟CBC會取最後五筆資料，時間=今天且最新的那筆，所以一過晚上12點就會撈不到
*   ABG取最新一筆，不管時間
*   Vital Sign中，體溫跟BP/HR/RR時間不一樣，會以BP/HR/RR的時間為主(因為護師量的時間不同)，因為後者量得比較頻繁。
*   IO是撈查房摘要裡的文字，會在早上七點後更新，所以你如果是在六點五十分產生的話會是前一天的舊資料(到底是那個天才要把IO藏在NIS裡，有夠難撈)
