Windows Scripting Host (WSH)

Windows Scripting Host — bu Windows operatsion tizimida mavjud bo‘lgan, batch fayllarni ishga tushirish orqali tizimdagi vazifalarni avtomatlashtirish va boshqarish uchun ishlatiladigan o‘rnatilgan boshqaruv vositasidir.

Bu Windows’ning o‘zida mavjud bo‘lgan engine bo‘lib, **cscript.exe** (komanda qatori skriptlari uchun) va **wscript.exe** (grafik interfeys skriptlari uchun) yordamida ishlaydi. Ular turli xil Microsoft Visual Basic Scripts (VBScript) kodlarini, jumladan **.vbs** va **.vbe** fayllarini bajarishga mas’uldir. VBScript haqida ko‘proq ma’lumot olish uchun shu yerga qarang.

Muhim jihati shundaki, Windows operatsion tizimidagi VBScript engine skriptlarni oddiy foydalanuvchi huquqlari va ruxsatlari bilan ishga tushiradi; shuning uchun u red teamerlar uchun foydali bo‘lishi mumkin.

Endi esa oddiy VBScript kodini yozamiz — bu kod Windows oynasida *"Welcome to THM"* degan xabarni chiqaradi. Quyidagi kodni masalan **hello.vbs** nomi bilan saqlang:

```vbscript
Dim message
message = "Welcome to THM"
MsgBox message
```

Birinchi qatorda **Dim** orqali `message` o‘zgaruvchisini e’lon qildik. Keyin unga `"Welcome to THM"` satr qiymatini berdik. Keyingi qatorda **MsgBox** funksiyasi yordamida o‘zgaruvchi tarkibini ekranga chiqaramiz. **MsgBox** funksiyasi haqida ko‘proq ma’lumot olish uchun shu yerga qarang.

Shundan so‘ng, **wscript** yordamida **hello.vbs** faylini ishga tushiramiz. Natijada ekranda *"Welcome to THM"* degan yozuvli Windows xabar oynasi paydo bo‘ladi.

![](https://raw.githubusercontent.com/akhatkulov/weaponization/refs/heads/main/Windows%20Scripting%20Host/Windows%20Scripting%20Host.png)

Keling, endi VBScript yordamida **.exe** fayllarni ishga tushiramiz. Quyidagi **vbs** kodi Windows kalkulyatorini ishga tushiradi va shu orqali Windows’ning o‘zida mavjud bo‘lgan engine (**WSH**) yordamida `.exe` fayllarni ishga tushirish mumkinligini isbotlaymiz:

```vbscript
Set shell = WScript.CreateObject("Wscript.Shell")
shell.Run("C:\Windows\System32\calc.exe " & WScript.ScriptFullName),0,True
```

Bu yerda biz **CreateObject** orqali **WScript** kutubxonasidan obyekt yaratib, bajariladigan faylni ishga tushirish (execution payload) chaqiramiz. Keyin **Run** metodi orqali ushbu faylni ishga tushiramiz. Ushbu misolda biz Windows kalkulyatori — **calc.exe** dasturini ishga tushiramiz.

**vbs** faylini ishga tushirish uchun **wscript** dan foydalanish mumkin:

```
c:\Windows\System32>wscript c:\Users\thm\Desktop\payload.vbs
```

Yoki **cscript** orqali ham ishga tushirish mumkin:

```
c:\Windows\System32>cscript.exe c:\Users\thm\Desktop\payload.vbs
```

Natijada, Windows kalkulyatori ishga tushadi va Desktop’da paydo bo‘ladi.

---

**Yana bir fokus:** agar `.vbs` fayllar bloklangan (blacklist) bo‘lsa, uni oddiygina `.txt` faylga qayta nomlab, quyidagicha ishga tushirish mumkin:

```
c:\Windows\System32>wscript /e:VBScript c:\Users\thm\Desktop\payload.txt
```
