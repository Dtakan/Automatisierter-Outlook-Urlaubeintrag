# Automatisierter-Outlook-Urlaubeintrag

Dieses Skript automatisiert das Hinzufügen von Urlaubsdaten zu Ihrem Outlook-Kalender. Es benötigt sowohl Outlook als auch die Urlaub-Seite geöffnet, um korrekt funktionieren zu können. Die Daten werden als Urlaub in Ihren Outlook-Kalender hinzugefügt.

Skript-Ablauf: Nachname eingeben, Dateipfad für .txt Datei angeben, Skript starten, zur Seite mit Urlaubsdaten wechseln, das Skript drückt die Tasten STRG + A und STRG + C und fügt den Inhalt in eine .txt Datei ein. Die .txt Datei wird ausgelesen und die relevanten Daten werden in Ihr Outlook Kalender eingetragen mit dem Betreff "Urlaub IhrName".

For programmers: ich habe nicht beautifulsoup verwendet, damit User ohne Python das Skript auch verwenden können ( + BS ist etwas wonky am Geschäfts-PC). Deshalb der workaround mit der .txt Datei. Weniger elegant aber if it looks stupid but it works, then it's not stupid.  
