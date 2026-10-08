/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 *
 * Elixir client PDF: strings of the progress text, the fallback message and the Options row, in the 12
 * languages of the editor's translated chrome. {0}, {1} are placeholders.
 */

define([], function () {
    'use strict';

    var CP = (window.Common = window.Common || {}).ElixirClientPdf = window.Common.ElixirClientPdf || {};

    CP.Strings = {
        en: {
            progress: 'Creating the PDF… page {0} of {1}',
            fallback: 'The PDF was created by the server because {0}. Fonts installed only on this computer are replaced by the server fonts in this file.',
            forms: 'the document contains form fields', restricted: 'the licence of the font {0} forbids embedding it',
            font: 'the font {0} could not be embedded', image: 'an image could not be read', size: 'the document is very long',
            other: 'some content is not supported yet by the browser export',
            ok: 'OK', group: 'PDF', option: 'Create PDF files in the browser',
            hint: 'Download as PDF and Print use the fonts shown on screen, including the fonts of this computer; the document is not sent to the server. Off: the server creates the PDF. PDF/A is always created by the server.'
        },
        fr: {
            progress: 'Création du PDF… page {0} sur {1}',
            fallback: 'Le PDF a été créé par le serveur car {0}. Dans ce fichier, les polices installées uniquement sur cet ordinateur sont remplacées par celles du serveur.',
            forms: 'le document contient des champs de formulaire', restricted: 'la licence de la police {0} interdit de l’incorporer',
            font: 'la police {0} n’a pas pu être incorporée', image: 'une image n’a pas pu être lue', size: 'le document est très long',
            other: 'certains contenus ne sont pas encore pris en charge par l’export dans le navigateur',
            ok: 'OK', group: 'PDF', option: 'Créer les fichiers PDF dans le navigateur',
            hint: 'Télécharger en PDF et Imprimer utilisent les polices affichées à l’écran, y compris celles de cet ordinateur ; le document n’est pas envoyé au serveur. Désactivé : le serveur crée le PDF. Le PDF/A est toujours créé par le serveur.'
        },
        de: {
            progress: 'PDF wird erstellt… Seite {0} von {1}',
            fallback: 'Die PDF-Datei wurde vom Server erstellt, weil {0}. Schriftarten, die nur auf diesem Computer installiert sind, werden in dieser Datei durch Serverschriftarten ersetzt.',
            forms: 'das Dokument Formularfelder enthält', restricted: 'die Lizenz der Schriftart {0} das Einbetten verbietet',
            font: 'die Schriftart {0} nicht eingebettet werden konnte', image: 'ein Bild nicht gelesen werden konnte', size: 'das Dokument sehr lang ist',
            other: 'einige Inhalte vom Export im Browser noch nicht unterstützt werden',
            ok: 'OK', group: 'PDF', option: 'PDF-Dateien im Browser erstellen',
            hint: 'Als PDF herunterladen und Drucken verwenden die am Bildschirm angezeigten Schriftarten, auch die dieses Computers; das Dokument wird nicht an den Server gesendet. Aus: Der Server erstellt die PDF-Datei. PDF/A wird immer vom Server erstellt.'
        },
        es: {
            progress: 'Creando el PDF… página {0} de {1}',
            fallback: 'El servidor creó el PDF porque {0}. En este archivo, las fuentes instaladas solo en este equipo se sustituyen por las del servidor.',
            forms: 'el documento contiene campos de formulario', restricted: 'la licencia de la fuente {0} prohíbe incrustarla',
            font: 'no se pudo incrustar la fuente {0}', image: 'no se pudo leer una imagen', size: 'el documento es muy largo',
            other: 'parte del contenido aún no es compatible con la exportación en el navegador',
            ok: 'Aceptar', group: 'PDF', option: 'Crear los archivos PDF en el navegador',
            hint: 'Descargar como PDF e Imprimir usan las fuentes que se ven en pantalla, incluidas las de este equipo; el documento no se envía al servidor. Desactivado: el servidor crea el PDF. El PDF/A siempre lo crea el servidor.'
        },
        it: {
            progress: 'Creazione del PDF… pagina {0} di {1}',
            fallback: 'Il PDF è stato creato dal server perché {0}. In questo file i caratteri installati solo su questo computer sono sostituiti da quelli del server.',
            forms: 'il documento contiene campi modulo', restricted: 'la licenza del carattere {0} ne vieta l’incorporamento',
            font: 'non è stato possibile incorporare il carattere {0}', image: 'non è stato possibile leggere un’immagine', size: 'il documento è molto lungo',
            other: 'alcuni contenuti non sono ancora supportati dall’esportazione nel browser',
            ok: 'OK', group: 'PDF', option: 'Crea i file PDF nel browser',
            hint: 'Scarica come PDF e Stampa usano i caratteri visualizzati sullo schermo, inclusi quelli di questo computer; il documento non viene inviato al server. Disattivato: il PDF viene creato dal server. Il PDF/A viene sempre creato dal server.'
        },
        pt: {
            progress: 'Criando o PDF… página {0} de {1}',
            fallback: 'O PDF foi criado pelo servidor porque {0}. Neste arquivo, as fontes instaladas apenas neste computador são substituídas pelas do servidor.',
            forms: 'o documento contém campos de formulário', restricted: 'a licença da fonte {0} proíbe incorporá-la',
            font: 'não foi possível incorporar a fonte {0}', image: 'não foi possível ler uma imagem', size: 'o documento é muito longo',
            other: 'parte do conteúdo ainda não é suportada pela exportação no navegador',
            ok: 'OK', group: 'PDF', option: 'Criar arquivos PDF no navegador',
            hint: 'Baixar como PDF e Imprimir usam as fontes exibidas na tela, inclusive as deste computador; o documento não é enviado ao servidor. Desativado: o servidor cria o PDF. O PDF/A é sempre criado pelo servidor.'
        },
        'pt-pt': {
            progress: 'A criar o PDF… página {0} de {1}',
            fallback: 'O PDF foi criado pelo servidor porque {0}. Neste ficheiro, os tipos de letra instalados apenas neste computador são substituídos pelos do servidor.',
            forms: 'o documento contém campos de formulário', restricted: 'a licença do tipo de letra {0} proíbe incorporá-lo',
            font: 'não foi possível incorporar o tipo de letra {0}', image: 'não foi possível ler uma imagem', size: 'o documento é muito longo',
            other: 'algum conteúdo ainda não é suportado pela exportação no navegador',
            ok: 'OK', group: 'PDF', option: 'Criar ficheiros PDF no navegador',
            hint: 'Transferir como PDF e Imprimir usam os tipos de letra mostrados no ecrã, incluindo os deste computador; o documento não é enviado para o servidor. Desativado: o servidor cria o PDF. O PDF/A é sempre criado pelo servidor.'
        },
        nl: {
            progress: 'PDF maken… pagina {0} van {1}',
            fallback: 'De pdf is door de server gemaakt omdat {0}. Lettertypen die alleen op deze computer staan, zijn in dit bestand vervangen door die van de server.',
            forms: 'het document formuliervelden bevat', restricted: 'de licentie van het lettertype {0} insluiten verbiedt',
            font: 'het lettertype {0} niet kon worden ingesloten', image: 'een afbeelding niet kon worden gelezen', size: 'het document erg lang is',
            other: 'sommige inhoud nog niet wordt ondersteund door de export in de browser',
            ok: 'OK', group: 'PDF', option: 'Pdf-bestanden in de browser maken',
            hint: 'Downloaden als pdf en Afdrukken gebruiken de lettertypen op het scherm, ook die van deze computer; het document wordt niet naar de server gestuurd. Uit: de server maakt de pdf. PDF/A wordt altijd door de server gemaakt.'
        },
        ru: {
            progress: 'Создание PDF… страница {0} из {1}',
            fallback: 'PDF создан сервером, потому что {0}. В этом файле шрифты, установленные только на этом компьютере, заменены шрифтами сервера.',
            forms: 'документ содержит поля формы', restricted: 'лицензия шрифта {0} запрещает его внедрение',
            font: 'не удалось внедрить шрифт {0}', image: 'не удалось прочитать изображение', size: 'документ очень длинный',
            other: 'часть содержимого пока не поддерживается экспортом в браузере',
            ok: 'ОК', group: 'PDF', option: 'Создавать PDF-файлы в браузере',
            hint: '«Скачать как PDF» и «Печать» используют шрифты, видимые на экране, включая шрифты этого компьютера; документ не отправляется на сервер. Выключено: PDF создаёт сервер. PDF/A всегда создаётся сервером.'
        },
        pl: {
            progress: 'Tworzenie pliku PDF… strona {0} z {1}',
            fallback: 'Plik PDF został utworzony przez serwer, ponieważ {0}. W tym pliku czcionki zainstalowane tylko na tym komputerze zastąpiono czcionkami serwera.',
            forms: 'dokument zawiera pola formularza', restricted: 'licencja czcionki {0} zabrania jej osadzania',
            font: 'nie udało się osadzić czcionki {0}', image: 'nie udało się odczytać obrazu', size: 'dokument jest bardzo długi',
            other: 'część zawartości nie jest jeszcze obsługiwana przez eksport w przeglądarce',
            ok: 'OK', group: 'PDF', option: 'Twórz pliki PDF w przeglądarce',
            hint: 'Pobierz jako PDF i Drukuj używają czcionek widocznych na ekranie, także czcionek tego komputera; dokument nie jest wysyłany na serwer. Wyłączone: plik PDF tworzy serwer. PDF/A zawsze tworzy serwer.'
        },
        ja: {
            progress: 'PDF を作成しています… {1} ページ中 {0} ページ',
            fallback: '{0}ため、PDF はサーバーで作成されました。このファイルでは、このコンピューターにのみインストールされているフォントはサーバーのフォントに置き換えられています。',
            forms: '文書にフォーム フィールドが含まれている', restricted: 'フォント {0} のライセンスで埋め込みが禁止されている',
            font: 'フォント {0} を埋め込めなかった', image: '画像を読み込めなかった', size: '文書が非常に長い',
            other: '一部の内容がブラウザーでのエクスポートにまだ対応していない',
            ok: 'OK', group: 'PDF', option: 'PDF ファイルをブラウザーで作成する',
            hint: '「PDF としてダウンロード」と「印刷」では、このコンピューターのフォントを含め、画面に表示されているフォントが使われます。文書はサーバーに送信されません。オフ: PDF はサーバーで作成されます。PDF/A は常にサーバーで作成されます。'
        },
        zh: {
            progress: '正在创建 PDF… 第 {0} 页，共 {1} 页',
            fallback: '由于{0}，PDF 由服务器创建。在此文件中，仅安装在这台电脑上的字体已替换为服务器字体。',
            forms: '文档包含表单域', restricted: '字体 {0} 的许可禁止嵌入',
            font: '无法嵌入字体 {0}', image: '无法读取某个图片', size: '文档过长',
            other: '部分内容暂不支持在浏览器中导出',
            ok: '确定', group: 'PDF', option: '在浏览器中创建 PDF 文件',
            hint: '“下载为 PDF”和“打印”使用屏幕上显示的字体，包括这台电脑上的字体；文档不会发送到服务器。关闭：由服务器创建 PDF。PDF/A 始终由服务器创建。'
        }
    };

    var lang = 'en';
    CP.setLang = function (l) {
        l = String(l || 'en').toLowerCase().replace('_', '-');
        lang = CP.Strings[l] ? l : (CP.Strings[l.split('-')[0]] ? l.split('-')[0] : 'en');
    };
    CP.t = function (key) {
        var s = (CP.Strings[lang] && CP.Strings[lang][key]) || CP.Strings.en[key] || key, args = arguments;
        return s.replace(/\{(\d)\}/g, function (m, i) { return args[+i + 1] !== undefined ? args[+i + 1] : ''; });
    };
    // engine fallback codes (sdkjs common/clientpdf) -> sentence part
    CP.reason = function (code, detail) {
        var font = String(detail || '').split(':')[0];
        switch (code) {
            case 'form-fields': return CP.t('forms');
            case 'font-restricted': return CP.t('restricted', font);
            case 'font-missing': case 'font-format': case 'font-subset': return CP.t('font', font);
            case 'image': return CP.t('image');
            case 'size': return CP.t('size');
            default: return CP.t('other');
        }
    };
    return CP;
});
