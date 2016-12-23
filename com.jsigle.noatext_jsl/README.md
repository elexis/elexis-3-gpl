Dieser Fork ist eine Portierung des Plugins auf Elexis 3.1, da OpenSource Elexis kein funktionierendes Textplugin mehr hat.
Keine externen Abhängigkeiten.

gweirich, 4.2.2016
___________________

NOAText_jsl
===========


Ein TextPlugin für Elexis, aktualisiert auf Basis von noa-libre, mit verbesserter Stabilität.

Version 1.4.5 benötigt die separate Datei ag.ion.noa_2.2.3.jar nicht mehr,
diese war in 1.4.3 noch erforderlich. Ausserdem wurde der Name des Plugins geändert:
von ch.elexis.noatext_jsl nach com.jsigle.noatext_jsl.

Ausführliche Informationen auf: http://www.jsigle.com/prog/elexis#NOAText_jsl

Status: BETA.
        Keinerlei Gewährleistung!
        Verwendung ausschliesslich auf eigene Verantwortung!
        Sachkundiger Anwender, eigene Funktionstests in
        unkritischer Umgebung und gute Backups dringend empfohlen!

Lizenz: Vorläufig GPL Version 2.1

https://github.com/LibreOffice/noa-libre

com.jsigle.noatext_jsl/src/com/jsigle/noa/NOAText.java
com.jsigle.noatext_jsl/src/ag/ion/bion/workbench/office/editor/core/EditorCorePlugin.java
com.jsigle.noatext_jsl/src/ag/ion/bion/officelayer/internal/document/DocumentLoader.java 
com.jsigle.noatext_jsl/src/ag/ion/noa4e/ui/wizards/application/LocalApplicationWizard.java
com.jsigle.noatext_jsl/src/ag/ion/noa4e/ui/NOAUIPlugin.java
com.jsigle.noatext_jsl/src/ag/ion/noa4e/ui/operations/LoadDocumentOperation.java
com.jsigle.noatext_jsl/src/ag/ion/noa4e/ui/widgets/OfficePanel.java
com.jsigle.noatext_jsl/src/ag/ion/noa4e/internal/ui/preferences/localOfficeApplicationPreferencesPage.java
com.jsigle.noatext_jsl/src/ag/ion/noa4e/internal/ui/preferences/LocalOfficeApplicationPreferencesPage.java
com.jsigle.noatext_jsl/src/ag/ion/noa4e/internal/ui/preferences/messages.properties
