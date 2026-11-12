إضافة Outlook Classic: نص افتتاحي حسب المستلم

مهم: هذه حزمة نموذجية غير مُختبرة على حساب Outlook الخاص بك. تحتاج حساب Microsoft 365 / Exchange Online / Outlook.com يدعم إضافات Office، واتصال إنترنت. إضافات Office.js لا تعمل على Gmail أو الحسابات المضافة عبر IMAP/POP داخل Outlook Classic.

1) تأكد من نوع الحساب:
Outlook > File > Account Settings > Account Settings > Email > Type

2) خصص البريد والجملة:
افتح ملف handlers.js في Notepad أو VS Code وعدّل قائمة RECIPIENT_RULES في بداية الملف.
يمكنك إضافة أو حذف قواعد. الصياغة:
 { address: "person@company.com", introduction: "Hello,\nYour custom message here." },
يجب إبقاء الفواصل وعلامات الاقتباس بالشكل الصحيح.

3) استضافة ملفات الإضافة على HTTPS:
الأسهل للاختبار هو GitHub Pages (لكن الملفات عامة ومكشوفة للآخرين؛ لا تضع فيها معلومات عمل سرية).
- أنشئ مستودعاً (Repository) عاماً باسم outlook-recipient-intro على GitHub.
- ارفع الملفات: index.html، runtime.html، handlers.js، manifest.xml ومجلد assets.
- افتح Settings > Pages، واختر Deploy from a branch > main > /(root)، ثم Save.
- افتح https://GITHUB_USERNAME.github.io/outlook-recipient-intro/ للتأكد أن الموقع منشور.
- لو أعددت الاستضافة على خادم مؤسسي بدلاً من GitHub، غيّر روابط URL في ملف manifest.xml لتطابق خادم HTTPS.

4) إعداد manifest.xml:
باستخدام Notepad، استبدل YOURUSERNAME في كل مكان باسم حساب GitHub الفعلي (بدون @)، واحفظ الملف.
إذا كان اسم المستودع مختلفاً عن outlook-recipient-intro يجب تعديل جزء المسار في كل الروابط أيضاً.
لا تعِد استخدام نفس معرف (Id) للإضافة مع إضافة أخرى.

5) التثبيت:
- افتح Outlook Classic واختر File > Info > Manage Add-ins.
- أو افتح https://aka.ms/olksideload من المتصفح المسجّل بنفس حساب Outlook.
- اختر My add-ins > Custom Addins > Add a custom add-in > Add from File.
- اختر ملف manifest.xml المعدّل على جهازك ووافِق على التثبيت.
- قد تتطلب المؤسسة موافقة مسؤول Microsoft 365، وقد يتأخر ظهور الإضافة في Outlook Classic بسبب التخزين المؤقت.

6) الاختبار:
- افتح رسالة جديدة داخل Outlook Classic.
- اكتب العنوان المستهدف في خانة To واضغط Enter أو Tab كي يتعرف Outlook على العنوان.
- يجب أن يظهر النص المحدد في بداية الرسالة، إذا كان الحدث مدعوماً وتعمل الإضافة بصورة صحيحة.
- لا تجرّب على رسالة مهمة قبل اختبار إرسال إلى حسابك الشخصي.

حدود هذه النسخة التجريبية:
- تطابق العنوان في خانة To فقط، وليس Cc / Bcc.
- عند تغيير عنوان المستلم بعد إدراج النص، لن تحذف النسخة التجريبية النص السابق تلقائياً؛ احذفه بنفسك.
- إذا كان في الرسالة أكثر من مستلم له قاعدة، تُطبّق أول قاعدة في RECIPIENT_RULES.
- لا تضيف النص مرتين إذا كان النص نفسه موجوداً فعلاً في الجسم.
- للاستخدام المؤسسي مع عناوين أو نصوص سرية استخدم خادم HTTPS خاصاً تديره الشركة لا GitHub Pages العام.
- هذه تجربة مبدئية لم يتم اختبارها على بيئتك أو ضمان توافقها مع سياسات شركتك.

مراجع Microsoft:
https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation
https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing
https://learn.microsoft.com/en-us/javascript/api/outlook/office.body?view=office-js
