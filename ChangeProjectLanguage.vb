' ****************************************************
' *                                                  *
' *       Welcome to My VB.NET Program               *
' *                                                  *
' *   Script written by: Davy Kyndt - KDPlan         *
' *                                                  *
' ****************************************************
'
'   ██╗  ██╗██████╗ ██████╗ ██╗      █████╗ ███╗   ██╗
'   ██║ ██╔╝██╔══██╗██╔══██╗██║     ██╔══██╗████╗  ██║
'   █████╔╝ ██║  ██║██████╔╝██║     ███████║██╔██╗ ██║
'   ██╔═██╗ ██║  ██║██╔═══╝ ██║     ██╔══██║██║╚██╗██║
'   ██║  ██╗██████╔╝██║     ███████╗██║  ██║██║ ╚████║
'   ╚═╝  ╚═╝╚═════╝ ╚═╝     ╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝
'
' ****************************************************
' Author : Davy Kyndt
' Project: Change the project language in EPLAN Electric P8
' 2 Ribbon tabs are created
' 1 button is created for each language with language flag, and 1 button for each dual language combination
' ****************************************************

Public Class ChangeProjectLanguage

#Region "Ribbon tab constants"
    ' =========================================================================
    ' Ribbon tab translations
    ' =========================================================================
    ' Define the tab title for each supported EPLAN language.
    ' Leave a value empty ("") to fall back to the default language.
    ' =========================================================================

    ' =========================================================================
    ' Ribbon configuration
    ' =========================================================================

    ' Ribbon tab name shown in EPLAN.
    Private Const TabName1 As String = "Project Language Single"
    Private Const TabName2 As String = "Project Language Double"


    ' Position of the tab within the ribbon.
    ' Higher values move the tab further to the right.
    Private Const TabOrder1 As Integer = 1100
    Private Const TabOrder2 As Integer = 1101

    ' Name of the command group inside the ribbon tab.
    Private Const GroupName1 As String = "Single Language"
    Private Const GroupName2 As String = "Double Language"

    ' Action called when the button is clicked.
    ' Must match the [DeclareAction] name.
    Private Const ActionName_de_DE As String = "German"
    Private Const ActionName_en_US As String = "English"
    Private Const ActionName_es_ES As String = "Spanish"
    Private Const ActionName_fr_FR As String = "French"
    Private Const ActionName_nl_NL As String = "Dutch"
    Private Const ActionName_sv_SE As String = "Swedish"
    Private Const ActionName_da_DK As String = "Danish"
    Private Const ActionName_ru_RU As String = "Russian"
    Private Const ActionName_zh_CN As String = "Chinese"
    Private Const ActionName_pl_PL As String = "Polish"
    Private Const ActionName_pt_BR As String = "Portuguese_Brazil"
    Private Const ActionName_cs_CZ As String = "Czech"
    Private Const ActionName_it_IT As String = "Italian"
    Private Const ActionName_hu_HU As String = "Hungarian"
    Private Const ActionName_pt_PT As String = "Portuguese_Portugal"
    Private Const ActionName_ko_KR As String = "Korean"
    Private Const ActionName_ja_JP As String = "Japanese"
    Private Const ActionName_tr_TR As String = "Turkish"
    Private Const ActionName_ro_RO As String = "Romanian"

    ' Button caption displayed on the ribbon.
    Private Const ButtonText_de_DE As String = "German"
    Private Const ButtonText_en_US As String = "English"
    Private Const ButtonText_es_ES As String = "Spanish"
    Private Const ButtonText_fr_FR As String = "French"
    Private Const ButtonText_nl_NL As String = "Dutch"
    Private Const ButtonText_sv_SE As String = "Swedish"
    Private Const ButtonText_da_DK As String = "Danish"
    Private Const ButtonText_ru_RU As String = "Russian"
    Private Const ButtonText_zh_CN As String = "Chinese"
    Private Const ButtonText_pl_PL As String = "Polish"
    Private Const ButtonText_pt_BR As String = "Portuguese Brazil"
    Private Const ButtonText_cs_CZ As String = "Czech"
    Private Const ButtonText_it_IT As String = "Italian"
    Private Const ButtonText_hu_HU As String = "Hungarian"
    Private Const ButtonText_pt_PT As String = "Portuguese Portugal"
    Private Const ButtonText_ko_KR As String = "Korean"
    Private Const ButtonText_ja_JP As String = "Japanese"
    Private Const ButtonText_tr_TR As String = "Turkish"
    Private Const ButtonText_ro_RO As String = "Romanian"

    ' Short description shown by EPLAN.
    Private Const ButtonDescription_de_DE As String = "Stellt die Projektsprache auf Deutsch ein"
    Private Const ButtonDescription_en_US As String = "Sets the project language to English"
    Private Const ButtonDescription_es_ES As String = "Establece el idioma del proyecto en español"
    Private Const ButtonDescription_fr_FR As String = "Définit la langue du projet sur le français"
    Private Const ButtonDescription_nl_NL As String = "Stelt de projecttaal in op Nederlands"
    Private Const ButtonDescription_sv_SE As String = "Ställer in projektspråket på svenska"
    Private Const ButtonDescription_da_DK As String = "Indstiller projektsproget til dansk"
    Private Const ButtonDescription_ru_RU As String = "Устанавливает язык проекта на русский"
    Private Const ButtonDescription_zh_CN As String = "将项目语言设置为中文"
    Private Const ButtonDescription_pl_PL As String = "Ustawia język projektu na polski"
    Private Const ButtonDescription_pt_BR As String = "Define o idioma do projeto para português do Brasil"
    Private Const ButtonDescription_cs_CZ As String = "Nastaví jazyk projektu na češtinu"
    Private Const ButtonDescription_it_IT As String = "Imposta la lingua del progetto su italiano"
    Private Const ButtonDescription_hu_HU As String = "Beállítja a projekt nyelvét magyarra"
    Private Const ButtonDescription_pt_PT As String = "Define o idioma do projeto para português de Portugal"
    Private Const ButtonDescription_ko_KR As String = "프로젝트 언어를 한국어로 설정합니다"
    Private Const ButtonDescription_ja_JP As String = "プロジェクトの言語を日本語に設定します"
    Private Const ButtonDescription_tr_TR As String = "Proje dilini Türkçe olarak ayarlar"
    Private Const ButtonDescription_ro_RO As String = "Setează limba proiectului pe română"

    ' Detailed tooltip shown when hovering over the button.
    Private Const ButtonTooltip_de_DE As String = "Klicken Sie, um die Projektsprache auf Deutsch zu ändern"
    Private Const ButtonTooltip_en_US As String = "Click to change the project language to English"
    Private Const ButtonTooltip_es_ES As String = "Haga clic para cambiar el idioma del proyecto a español"
    Private Const ButtonTooltip_fr_FR As String = "Cliquez pour changer la langue du projet en français"
    Private Const ButtonTooltip_nl_NL As String = "Klik om de projecttaal te wijzigen in het Nederlands"
    Private Const ButtonTooltip_sv_SE As String = "Klicka för att ändra projektspråket till svenska"
    Private Const ButtonTooltip_da_DK As String = "Klik for at ændre projektsproget til dansk"
    Private Const ButtonTooltip_ru_RU As String = "Нажмите, чтобы изменить язык проекта на русский"
    Private Const ButtonTooltip_zh_CN As String = "单击以将项目语言更改为中文"
    Private Const ButtonTooltip_pl_PL As String = "Kliknij, aby zmienić język projektu na polski"
    Private Const ButtonTooltip_pt_BR As String = "Clique para alterar o idioma do projeto para português do Brasil"
    Private Const ButtonTooltip_cs_CZ As String = "Klikněte pro změnu jazyka projektu na češtinu"
    Private Const ButtonTooltip_it_IT As String = "Fare clic per cambiare la lingua del progetto in italiano"
    Private Const ButtonTooltip_hu_HU As String = "Kattintson a projekt nyelvének magyarra történő módosításához"
    Private Const ButtonTooltip_pt_PT As String = "Clique para alterar o idioma do projeto para português de Portugal"
    Private Const ButtonTooltip_ko_KR As String = "클릭하여 프로젝트 언어를 한국어로 변경합니다"
    Private Const ButtonTooltip_ja_JP As String = "クリックしてプロジェクトの言語を日本語に変更します"
    Private Const ButtonTooltip_tr_TR As String = "Proje dilini Türkçe olarak değiştirmek için tıklayın"
    Private Const ButtonTooltip_ro_RO As String = "Faceți clic pentru a schimba limba proiectului în română"

    ' Dual language button configuration (optional)
    ' ============================================================
' 2-language project language ribbon actions
' ============================================================

    Private Const ActionName_nl_en As String = "Dutch_English"
    Private Const ButtonText_nl_en As String = "Dutch / English"
    Private Const ButtonDescription_nl_en As String = "Sets the project language to Dutch and English"
    Private Const ButtonTooltip_nl_en As String = "Click to change the project language to Dutch and English"

    Private Const ActionName_nl_fr As String = "Dutch_French"
    Private Const ButtonText_nl_fr As String = "Dutch / French"
    Private Const ButtonDescription_nl_fr As String = "Sets the project language to Dutch and French"
    Private Const ButtonTooltip_nl_fr As String = "Click to change the project language to Dutch and French"

    Private Const ActionName_nl_de As String = "Dutch_German"
    Private Const ButtonText_nl_de As String = "Dutch / German"
    Private Const ButtonDescription_nl_de As String = "Sets the project language to Dutch and German"
    Private Const ButtonTooltip_nl_de As String = "Click to change the project language to Dutch and German"

    Private Const ActionName_de_en As String = "German_English"
    Private Const ButtonText_de_en As String = "German / English"
    Private Const ButtonDescription_de_en As String = "Sets the project language to German and English"
    Private Const ButtonTooltip_de_en As String = "Click to change the project language to German and English"

    Private Const ActionName_de_fr As String = "German_French"
    Private Const ButtonText_de_fr As String = "German / French"
    Private Const ButtonDescription_de_fr As String = "Sets the project language to German and French"
    Private Const ButtonTooltip_de_fr As String = "Click to change the project language to German and French"

    Private Const ActionName_de_it As String = "German_Italian"
    Private Const ButtonText_de_it As String = "German / Italian"
    Private Const ButtonDescription_de_it As String = "Sets the project language to German and Italian"
    Private Const ButtonTooltip_de_it As String = "Click to change the project language to German and Italian"

    Private Const ActionName_de_pl As String = "German_Polish"
    Private Const ButtonText_de_pl As String = "German / Polish"
    Private Const ButtonDescription_de_pl As String = "Sets the project language to German and Polish"
    Private Const ButtonTooltip_de_pl As String = "Click to change the project language to German and Polish"

    Private Const ActionName_de_cz As String = "German_Czech"
    Private Const ButtonText_de_cz As String = "German / Czech"
    Private Const ButtonDescription_de_cz As String = "Sets the project language to German and Czech"
    Private Const ButtonTooltip_de_cz As String = "Click to change the project language to German and Czech"

    Private Const ActionName_de_hu As String = "German_Hungarian"
    Private Const ButtonText_de_hu As String = "German / Hungarian"
    Private Const ButtonDescription_de_hu As String = "Sets the project language to German and Hungarian"
    Private Const ButtonTooltip_de_hu As String = "Click to change the project language to German and Hungarian"

    Private Const ActionName_de_ro As String = "German_Romanian"
    Private Const ButtonText_de_ro As String = "German / Romanian"
    Private Const ButtonDescription_de_ro As String = "Sets the project language to German and Romanian"
    Private Const ButtonTooltip_de_ro As String = "Click to change the project language to German and Romanian"

    Private Const ActionName_de_ru As String = "German_Russian"
    Private Const ButtonText_de_ru As String = "German / Russian"
    Private Const ButtonDescription_de_ru As String = "Sets the project language to German and Russian"
    Private Const ButtonTooltip_de_ru As String = "Click to change the project language to German and Russian"

    Private Const ActionName_de_se As String = "German_Swedish"
    Private Const ButtonText_de_se As String = "German / Swedish"
    Private Const ButtonDescription_de_se As String = "Sets the project language to German and Swedish"
    Private Const ButtonTooltip_de_se As String = "Click to change the project language to German and Swedish"

    Private Const ActionName_de_da As String = "German_Danish"
    Private Const ButtonText_de_da As String = "German / Danish"
    Private Const ButtonDescription_de_da As String = "Sets the project language to German and Danish"
    Private Const ButtonTooltip_de_da As String = "Click to change the project language to German and Danish"

    Private Const ActionName_fr_en As String = "French_English"
    Private Const ButtonText_fr_en As String = "French / English"
    Private Const ButtonDescription_fr_en As String = "Sets the project language to French and English"
    Private Const ButtonTooltip_fr_en As String = "Click to change the project language to French and English"

    Private Const ActionName_fr_de As String = "French_German"
    Private Const ButtonText_fr_de As String = "French / German"
    Private Const ButtonDescription_fr_de As String = "Sets the project language to French and German"
    Private Const ButtonTooltip_fr_de As String = "Click to change the project language to French and German"

    Private Const ActionName_fr_nl As String = "French_Dutch"
    Private Const ButtonText_fr_nl As String = "French / Dutch"
    Private Const ButtonDescription_fr_nl As String = "Sets the project language to French and Dutch"
    Private Const ButtonTooltip_fr_nl As String = "Click to change the project language to French and Dutch"

    Private Const ActionName_fr_it As String = "French_Italian"
    Private Const ButtonText_fr_it As String = "French / Italian"
    Private Const ButtonDescription_fr_it As String = "Sets the project language to French and Italian"
    Private Const ButtonTooltip_fr_it As String = "Click to change the project language to French and Italian"

    Private Const ActionName_fr_es As String = "French_Spanish"
    Private Const ButtonText_fr_es As String = "French / Spanish"
    Private Const ButtonDescription_fr_es As String = "Sets the project language to French and Spanish"
    Private Const ButtonTooltip_fr_es As String = "Click to change the project language to French and Spanish"

    Private Const ActionName_es_en As String = "Spanish_English"
    Private Const ButtonText_es_en As String = "Spanish / English"
    Private Const ButtonDescription_es_en As String = "Sets the project language to Spanish and English"
    Private Const ButtonTooltip_es_en As String = "Click to change the project language to Spanish and English"

    Private Const ActionName_es_fr As String = "Spanish_French"
    Private Const ButtonText_es_fr As String = "Spanish / French"
    Private Const ButtonDescription_es_fr As String = "Sets the project language to Spanish and French"
    Private Const ButtonTooltip_es_fr As String = "Click to change the project language to Spanish and French"

    Private Const ActionName_es_pt As String = "Spanish_Portuguese"
    Private Const ButtonText_es_pt As String = "Spanish / Portuguese"
    Private Const ButtonDescription_es_pt As String = "Sets the project language to Spanish and Portuguese"
    Private Const ButtonTooltip_es_pt As String = "Click to change the project language to Spanish and Portuguese"

    Private Const ActionName_es_it As String = "Spanish_Italian"
    Private Const ButtonText_es_it As String = "Spanish / Italian"
    Private Const ButtonDescription_es_it As String = "Sets the project language to Spanish and Italian"
    Private Const ButtonTooltip_es_it As String = "Click to change the project language to Spanish and Italian"

    Private Const ActionName_it_en As String = "Italian_English"
    Private Const ButtonText_it_en As String = "Italian / English"
    Private Const ButtonDescription_it_en As String = "Sets the project language to Italian and English"
    Private Const ButtonTooltip_it_en As String = "Click to change the project language to Italian and English"

    Private Const ActionName_it_de As String = "Italian_German"
    Private Const ButtonText_it_de As String = "Italian / German"
    Private Const ButtonDescription_it_de As String = "Sets the project language to Italian and German"
    Private Const ButtonTooltip_it_de As String = "Click to change the project language to Italian and German"

    Private Const ActionName_it_fr As String = "Italian_French"
    Private Const ButtonText_it_fr As String = "Italian / French"
    Private Const ButtonDescription_it_fr As String = "Sets the project language to Italian and French"
    Private Const ButtonTooltip_it_fr As String = "Click to change the project language to Italian and French"

    Private Const ActionName_it_es As String = "Italian_Spanish"
    Private Const ButtonText_it_es As String = "Italian / Spanish"
    Private Const ButtonDescription_it_es As String = "Sets the project language to Italian and Spanish"
    Private Const ButtonTooltip_it_es As String = "Click to change the project language to Italian and Spanish"

    Private Const ActionName_pl_en As String = "Polish_English"
    Private Const ButtonText_pl_en As String = "Polish / English"
    Private Const ButtonDescription_pl_en As String = "Sets the project language to Polish and English"
    Private Const ButtonTooltip_pl_en As String = "Click to change the project language to Polish and English"

    Private Const ActionName_pl_de As String = "Polish_German"
    Private Const ButtonText_pl_de As String = "Polish / German"
    Private Const ButtonDescription_pl_de As String = "Sets the project language to Polish and German"
    Private Const ButtonTooltip_pl_de As String = "Click to change the project language to Polish and German"

    Private Const ActionName_pl_cz As String = "Polish_Czech"
    Private Const ButtonText_pl_cz As String = "Polish / Czech"
    Private Const ButtonDescription_pl_cz As String = "Sets the project language to Polish and Czech"
    Private Const ButtonTooltip_pl_cz As String = "Click to change the project language to Polish and Czech"

    Private Const ActionName_cz_en As String = "Czech_English"
    Private Const ButtonText_cz_en As String = "Czech / English"
    Private Const ButtonDescription_cz_en As String = "Sets the project language to Czech and English"
    Private Const ButtonTooltip_cz_en As String = "Click to change the project language to Czech and English"

    Private Const ActionName_cz_de As String = "Czech_German"
    Private Const ButtonText_cz_de As String = "Czech / German"
    Private Const ButtonDescription_cz_de As String = "Sets the project language to Czech and German"
    Private Const ButtonTooltip_cz_de As String = "Click to change the project language to Czech and German"

    Private Const ActionName_cz_pl As String = "Czech_Polish"
    Private Const ButtonText_cz_pl As String = "Czech / Polish"
    Private Const ButtonDescription_cz_pl As String = "Sets the project language to Czech and Polish"
    Private Const ButtonTooltip_cz_pl As String = "Click to change the project language to Czech and Polish"

    Private Const ActionName_cz_hu As String = "Czech_Hungarian"
    Private Const ButtonText_cz_hu As String = "Czech / Hungarian"
    Private Const ButtonDescription_cz_hu As String = "Sets the project language to Czech and Hungarian"
    Private Const ButtonTooltip_cz_hu As String = "Click to change the project language to Czech and Hungarian"

    Private Const ActionName_hu_en As String = "Hungarian_English"
    Private Const ButtonText_hu_en As String = "Hungarian / English"
    Private Const ButtonDescription_hu_en As String = "Sets the project language to Hungarian and English"
    Private Const ButtonTooltip_hu_en As String = "Click to change the project language to Hungarian and English"

    Private Const ActionName_hu_de As String = "Hungarian_German"
    Private Const ButtonText_hu_de As String = "Hungarian / German"
    Private Const ButtonDescription_hu_de As String = "Sets the project language to Hungarian and German"
    Private Const ButtonTooltip_hu_de As String = "Click to change the project language to Hungarian and German"

    Private Const ActionName_hu_cz As String = "Hungarian_Czech"
    Private Const ButtonText_hu_cz As String = "Hungarian / Czech"
    Private Const ButtonDescription_hu_cz As String = "Sets the project language to Hungarian and Czech"
    Private Const ButtonTooltip_hu_cz As String = "Click to change the project language to Hungarian and Czech"

    Private Const ActionName_hu_ro As String = "Hungarian_Romanian"
    Private Const ButtonText_hu_ro As String = "Hungarian / Romanian"
    Private Const ButtonDescription_hu_ro As String = "Sets the project language to Hungarian and Romanian"
    Private Const ButtonTooltip_hu_ro As String = "Click to change the project language to Hungarian and Romanian"

    Private Const ActionName_se_en As String = "Swedish_English"
    Private Const ButtonText_se_en As String = "Swedish / English"
    Private Const ButtonDescription_se_en As String = "Sets the project language to Swedish and English"
    Private Const ButtonTooltip_se_en As String = "Click to change the project language to Swedish and English"

    Private Const ActionName_se_da As String = "Swedish_Danish"
    Private Const ButtonText_se_da As String = "Swedish / Danish"
    Private Const ButtonDescription_se_da As String = "Sets the project language to Swedish and Danish"
    Private Const ButtonTooltip_se_da As String = "Click to change the project language to Swedish and Danish"

    Private Const ActionName_se_de As String = "Swedish_German"
    Private Const ButtonText_se_de As String = "Swedish / German"
    Private Const ButtonDescription_se_de As String = "Sets the project language to Swedish and German"
    Private Const ButtonTooltip_se_de As String = "Click to change the project language to Swedish and German"

    Private Const ActionName_da_en As String = "Danish_English"
    Private Const ButtonText_da_en As String = "Danish / English"
    Private Const ButtonDescription_da_en As String = "Sets the project language to Danish and English"
    Private Const ButtonTooltip_da_en As String = "Click to change the project language to Danish and English"

    Private Const ActionName_da_se As String = "Danish_Swedish"
    Private Const ButtonText_da_se As String = "Danish / Swedish"
    Private Const ButtonDescription_da_se As String = "Sets the project language to Danish and Swedish"
    Private Const ButtonTooltip_da_se As String = "Click to change the project language to Danish and Swedish"

    Private Const ActionName_da_de As String = "Danish_German"
    Private Const ButtonText_da_de As String = "Danish / German"
    Private Const ButtonDescription_da_de As String = "Sets the project language to Danish and German"
    Private Const ButtonTooltip_da_de As String = "Click to change the project language to Danish and German"

    Private Const ActionName_pt_en As String = "Portuguese_English"
    Private Const ButtonText_pt_en As String = "Portuguese / English"
    Private Const ButtonDescription_pt_en As String = "Sets the project language to Portuguese and English"
    Private Const ButtonTooltip_pt_en As String = "Click to change the project language to Portuguese and English"

    Private Const ActionName_pt_es As String = "Portuguese_Spanish"
    Private Const ButtonText_pt_es As String = "Portuguese / Spanish"
    Private Const ButtonDescription_pt_es As String = "Sets the project language to Portuguese and Spanish"
    Private Const ButtonTooltip_pt_es As String = "Click to change the project language to Portuguese and Spanish"

    Private Const ActionName_br_en As String = "Portuguese_Brazil_English"
    Private Const ButtonText_br_en As String = "Portuguese Brazil / English"
    Private Const ButtonDescription_br_en As String = "Sets the project language to Portuguese Brazil and English"
    Private Const ButtonTooltip_br_en As String = "Click to change the project language to Portuguese Brazil and English"

    Private Const ActionName_ro_en As String = "Romanian_English"
    Private Const ButtonText_ro_en As String = "Romanian / English"
    Private Const ButtonDescription_ro_en As String = "Sets the project language to Romanian and English"
    Private Const ButtonTooltip_ro_en As String = "Click to change the project language to Romanian and English"

    Private Const ActionName_ro_de As String = "Romanian_German"
    Private Const ButtonText_ro_de As String = "Romanian / German"
    Private Const ButtonDescription_ro_de As String = "Sets the project language to Romanian and German"
    Private Const ButtonTooltip_ro_de As String = "Click to change the project language to Romanian and German"

    Private Const ActionName_ro_hu As String = "Romanian_Hungarian"
    Private Const ButtonText_ro_hu As String = "Romanian / Hungarian"
    Private Const ButtonDescription_ro_hu As String = "Sets the project language to Romanian and Hungarian"
    Private Const ButtonTooltip_ro_hu As String = "Click to change the project language to Romanian and Hungarian"

    Private Const ActionName_tr_en As String = "Turkish_English"
    Private Const ButtonText_tr_en As String = "Turkish / English"
    Private Const ButtonDescription_tr_en As String = "Sets the project language to Turkish and English"
    Private Const ButtonTooltip_tr_en As String = "Click to change the project language to Turkish and English"

    Private Const ActionName_ru_en As String = "Russian_English"
    Private Const ButtonText_ru_en As String = "Russian / English"
    Private Const ButtonDescription_ru_en As String = "Sets the project language to Russian and English"
    Private Const ButtonTooltip_ru_en As String = "Click to change the project language to Russian and English"

    Private Const ActionName_zh_en As String = "Chinese_English"
    Private Const ButtonText_zh_en As String = "Chinese / English"
    Private Const ButtonDescription_zh_en As String = "Sets the project language to Chinese and English"
    Private Const ButtonTooltip_zh_en As String = "Click to change the project language to Chinese and English"

    Private Const ActionName_ja_en As String = "Japanese_English"
    Private Const ButtonText_ja_en As String = "Japanese / English"
    Private Const ButtonDescription_ja_en As String = "Sets the project language to Japanese and English"
    Private Const ButtonTooltip_ja_en As String = "Click to change the project language to Japanese and English"

    Private Const ActionName_ko_en As String = "Korean_English"
    Private Const ButtonText_ko_en As String = "Korean / English"
    Private Const ButtonDescription_ko_en As String = "Sets the project language to Korean and English"
    Private Const ButtonTooltip_ko_en As String = "Click to change the project language to Korean and English"

    Private Const ActionName_zh_ja As String = "Chinese_Japanese"
    Private Const ButtonText_zh_ja As String = "Chinese / Japanese"
    Private Const ButtonDescription_zh_ja As String = "Sets the project language to Chinese and Japanese"
    Private Const ButtonTooltip_zh_ja As String = "Click to change the project language to Chinese and Japanese"

    Private Const ActionName_zh_ko As String = "Chinese_Korean"
    Private Const ButtonText_zh_ko As String = "Chinese / Korean"
    Private Const ButtonDescription_zh_ko As String = "Sets the project language to Chinese and Korean"
    Private Const ButtonTooltip_zh_ko As String = "Click to change the project language to Chinese and Korean"

    Private Const ActionName_ja_ko As String = "Japanese_Korean"
    Private Const ButtonText_ja_ko As String = "Japanese / Korean"
    Private Const ButtonDescription_ja_ko As String = "Sets the project language to Japanese and Korean"
    Private Const ButtonTooltip_ja_ko As String = "Click to change the project language to Japanese and Korean"

    Private Const DefaultFolder As String = "D:\KDPlan\SynologyDrive\KDPlan Industry\Scripts\ChangeProjectLanguage\flags\"

    'Reserve CLI and ACC for later use in the script
    Private MyEplanSetting As New Eplan.EplApi.Base.Settings
    Private CLI As New Eplan.EplApi.ApplicationFramework.CommandLineInterpreter
    Private ACC As New Eplan.EplApi.ApplicationFramework.ActionCallingContext
#End Region

#Region "Register / Unregister"

    ''' <summary>
    ''' Creates the ribbon tab, command group and button when the script
    ''' is registered in EPLAN.
    ''' </summary>
    <Eplan.EplApi.Scripting.DeclareRegister>
    Public Sub RegisterRibbon()

        'single language icons
        Dim FlagIcon_AU As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_AU)
        Dim FlagIcon_BR As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_BR)
        Dim FlagIcon_CH As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_CH)
        Dim FlagIcon_CZ As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_CZ)
        Dim FlagIcon_DA As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_DA)
        Dim FlagIcon_DE As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_DE)
        Dim FlagIcon_EN As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_EN)
        Dim FlagIcon_ES As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_ES)
        Dim FlagIcon_FR As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_FR)
        Dim FlagIcon_HU As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_HU)
        Dim FlagIcon_IT As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_IT)
        Dim FlagIcon_JA As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_JA)
        Dim FlagIcon_KO As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_KO)
        Dim FlagIcon_NL As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_NL)
        Dim FlagIcon_PL As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_PL)
        Dim FlagIcon_PT As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_PT)
        Dim FlagIcon_RO As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_RO)
        Dim FlagIcon_RU As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_RU)
        Dim FlagIcon_SE As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_SE)
        Dim FlagIcon_TR As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_TR)
        Dim FlagIcon_US As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_US)
        Dim FlagIcon_ZH As new Eplan.EplApi.Gui.ribbonicon(Eplan.EplApi.Gui.CommandIcon.Flag_ZH)

        'dual language icons

        Dim FlagIcon_NL_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_NL_EN.svg")
        Dim FlagIcon_NL_FR As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_NL_FR.svg")
        Dim FlagIcon_NL_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Dutch-German_flag_hybrid.svg")

        Dim FlagIcon_DE_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_the_United_Kingdom_and_Germany.svg")
        Dim FlagIcon_DE_FR As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_France_and_Germany.svg")
        Dim FlagIcon_DE_IT As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_Italy_and_Germany.svg")
        Dim FlagIcon_DE_PL As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_Poland_and_Germany.svg")
        Dim FlagIcon_DE_CZ As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_the_Czech_Republic_and_Germany.svg")
        Dim FlagIcon_DE_HU As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_DE_HU.svg")
        Dim FlagIcon_DE_RO As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_Romania_and_Germany.svg")
        Dim FlagIcon_DE_RU As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_Russia_and_Germany.svg")
        Dim FlagIcon_DE_SE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_DE_SE.svg")
        Dim FlagIcon_DE_DA As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_DE_DA.svg")

        Dim FlagIcon_FR_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_France_and_the_United_Kingdom.svg")
        Dim FlagIcon_FR_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_France_and_Germany.svg")
        Dim FlagIcon_FR_NL As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_FR_NL.svg")
        Dim FlagIcon_FR_IT As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_FR_IT.svg")
        Dim FlagIcon_FR_ES As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_FR_ES.svg")

        Dim FlagIcon_ES_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_ES_EN.svg")
        Dim FlagIcon_ES_FR As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_ES_FR.svg")
        Dim FlagIcon_ES_PT As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_ES_PT.svg")
        Dim FlagIcon_ES_IT As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_ES_IT.svg")

        Dim FlagIcon_IT_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_the_United_States_and_Italy.svg")
        Dim FlagIcon_IT_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_Italy_and_Germany.svg")
        Dim FlagIcon_IT_FR As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_IT_FR.svg")
        Dim FlagIcon_IT_ES As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_IT_ES.svg")

        Dim FlagIcon_PL_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_PL_EN.svg")
        Dim FlagIcon_PL_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_Poland_and_Germany.svg")
        Dim FlagIcon_PL_CZ As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_PL_CZ.svg")

        Dim FlagIcon_CZ_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_CZ_EN.svg")
        Dim FlagIcon_CZ_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_the_Czech_Republic_and_Germany.svg")
        Dim FlagIcon_CZ_PL As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_CZ_PL.svg")
        Dim FlagIcon_CZ_HU As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_CZ_HU.svg")

        Dim FlagIcon_HU_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_HU_EN.svg")
        Dim FlagIcon_HU_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_HU_DE.svg")
        Dim FlagIcon_HU_CZ As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_HU_CZ.svg")
        Dim FlagIcon_HU_RO As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_HU_RO.svg")

        Dim FlagIcon_SE_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_SE_EN.svg")
        Dim FlagIcon_SE_DA As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_SE_DA.svg")
        Dim FlagIcon_SE_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_SE_DE.svg")

        Dim FlagIcon_DA_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_DA_EN.svg")
        Dim FlagIcon_DA_SE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_DA_SE.svg")
        Dim FlagIcon_DA_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_DA_DE.svg")

        Dim FlagIcon_PT_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_PT_EN.svg")
        Dim FlagIcon_PT_ES As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_PT_ES.svg")
        Dim FlagIcon_BR_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_BR_EN.svg")

        Dim FlagIcon_RO_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_RO_EN.svg")
        Dim FlagIcon_RO_DE As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "File_Flag_of_Romania_and_Germany.svg")
        Dim FlagIcon_RO_HU As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_RO_HU.svg")

        Dim FlagIcon_TR_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_TR_EN.svg")
        Dim FlagIcon_RU_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_RU_EN.svg")

        Dim FlagIcon_ZH_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_ZH_EN.svg")
        Dim FlagIcon_JA_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_JA_EN.svg")
        Dim FlagIcon_KO_EN As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_KO_EN.svg")
        Dim FlagIcon_ZH_JA As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_ZH_JA.svg")
        Dim FlagIcon_ZH_KO As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_ZH_KO.svg")
        Dim FlagIcon_JA_KO As New Eplan.EplApi.Gui.RibbonIcon(DefaultFolder & "Flag_JA_KO.svg")

        Dim ribbon As New Eplan.EplApi.Gui.RibbonBar()
        Dim ribbonTab1 As Eplan.EplApi.Gui.RibbonTab = Nothing
        Dim ribbonTab2 As Eplan.EplApi.Gui.RibbonTab = Nothing
        Dim tabTitle As New Eplan.EplApi.Base.MultiLangString()

        Dim TabName1 As String = "Project Language Single"
        Dim TabName2 As String = "Project Language Double"

        ' ---------------------------------------------------------------------
        ' Check if the ribbon tab already exists
        ' ---------------------------------------------------------------------

        For Each existingTab As Eplan.EplApi.Gui.RibbonTab In ribbon.Tabs
            If String.Equals(existingTab.Name, TabName1, StringComparison.OrdinalIgnoreCase) Then
                ribbonTab1 = existingTab
            End If
        Next

        ' Create the tab if it does not yet exist.
        If ribbonTab1 Is Nothing Then
            ribbonTab1 = ribbon.AddTab(TabName1, TabOrder1)
        End If


        For Each existingTab As Eplan.EplApi.Gui.RibbonTab In ribbon.Tabs
            If String.Equals(existingTab.Name, TabName2, StringComparison.OrdinalIgnoreCase) Then
                ribbonTab2 = existingTab
            End If
        Next

        If ribbonTab2 Is Nothing Then
            ribbonTab2 = ribbon.AddTab(TabName2, TabOrder2)
        End If

        If ribbonTab1 Is Nothing OrElse ribbonTab2 Is Nothing Then
            Exit Sub
        End If

        ' ---------------------------------------------------------------------
        ' Prevent duplicate command groups
        ' ---------------------------------------------------------------------

        For Each existingGroup As Eplan.EplApi.Gui.RibbonCommandGroup In ribbonTab1.CommandGroups
            If String.Equals(existingGroup.Name, GroupName1, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next

        For Each existingGroup As Eplan.EplApi.Gui.RibbonCommandGroup In ribbonTab2.CommandGroups
            If String.Equals(existingGroup.Name, GroupName2, StringComparison.OrdinalIgnoreCase) Then
                Exit Sub
            End If
        Next

        ' ---------------------------------------------------------------------
        ' Create command group and button
        ' ---------------------------------------------------------------------

        Dim commandGroup1 As Eplan.EplApi.Gui.RibbonCommandGroup = ribbonTab1.AddCommandGroup(GroupName1)


        commandGroup1.AddCommand(ButtonText_de_DE, ActionName_de_DE, ButtonDescription_de_DE, ButtonTooltip_de_DE, FlagIcon_DE)
        commandGroup1.AddCommand(ButtonText_en_US, ActionName_en_US, ButtonDescription_en_US, ButtonTooltip_en_US, FlagIcon_EN)
        commandGroup1.AddCommand(ButtonText_es_ES, ActionName_es_ES, ButtonDescription_es_ES, ButtonTooltip_es_ES, FlagIcon_ES)
        commandGroup1.AddCommand(ButtonText_fr_FR, ActionName_fr_FR, ButtonDescription_fr_FR, ButtonTooltip_fr_FR, FlagIcon_FR)
        commandGroup1.AddCommand(ButtonText_nl_NL, ActionName_nl_NL, ButtonDescription_nl_NL, ButtonTooltip_nl_NL, FlagIcon_NL)
        commandGroup1.AddCommand(ButtonText_sv_SE, ActionName_sv_SE, ButtonDescription_sv_SE, ButtonTooltip_sv_SE, FlagIcon_SE)
        commandGroup1.AddCommand(ButtonText_da_DK, ActionName_da_DK, ButtonDescription_da_DK, ButtonTooltip_da_DK, FlagIcon_DA)
        commandGroup1.AddCommand(ButtonText_ru_RU, ActionName_ru_RU, ButtonDescription_ru_RU, ButtonTooltip_ru_RU, FlagIcon_RU)
        commandGroup1.AddCommand(ButtonText_zh_CN, ActionName_zh_CN, ButtonDescription_zh_CN, ButtonTooltip_zh_CN, FlagIcon_ZH)
        commandGroup1.AddCommand(ButtonText_pl_PL, ActionName_pl_PL, ButtonDescription_pl_PL, ButtonTooltip_pl_PL, FlagIcon_PL)
        commandGroup1.AddCommand(ButtonText_pt_BR, ActionName_pt_BR, ButtonDescription_pt_BR, ButtonTooltip_pt_BR, FlagIcon_BR)
        commandGroup1.AddCommand(ButtonText_cs_CZ, ActionName_cs_CZ, ButtonDescription_cs_CZ, ButtonTooltip_cs_CZ, FlagIcon_CZ)
        commandGroup1.AddCommand(ButtonText_it_IT, ActionName_it_IT, ButtonDescription_it_IT, ButtonTooltip_it_IT, FlagIcon_IT)
        commandGroup1.AddCommand(ButtonText_hu_HU, ActionName_hu_HU, ButtonDescription_hu_HU, ButtonTooltip_hu_HU, FlagIcon_HU)
        commandGroup1.AddCommand(ButtonText_pt_PT, ActionName_pt_PT, ButtonDescription_pt_PT, ButtonTooltip_pt_PT, FlagIcon_PT)
        commandGroup1.AddCommand(ButtonText_ko_KR, ActionName_ko_KR, ButtonDescription_ko_KR, ButtonTooltip_ko_KR, FlagIcon_KO)
        commandGroup1.AddCommand(ButtonText_ja_JP, ActionName_ja_JP, ButtonDescription_ja_JP, ButtonTooltip_ja_JP, FlagIcon_JA)
        commandGroup1.AddCommand(ButtonText_tr_TR, ActionName_tr_TR, ButtonDescription_tr_TR, ButtonTooltip_tr_TR, FlagIcon_TR)
        commandGroup1.AddCommand(ButtonText_ro_RO, ActionName_ro_RO, ButtonDescription_ro_RO, ButtonTooltip_ro_RO, FlagIcon_RO)

        ' ============================================================
        ' Add commands to ribbon command group
        ' ============================================================
        Dim commandGroup2 As Eplan.EplApi.Gui.RibbonCommandGroup = ribbonTab2.AddCommandGroup(GroupName2)

        commandGroup2.AddCommand(ButtonText_nl_en, ActionName_nl_en, ButtonDescription_nl_en, ButtonTooltip_nl_en, FlagIcon_NL_EN)
        commandGroup2.AddCommand(ButtonText_nl_fr, ActionName_nl_fr, ButtonDescription_nl_fr, ButtonTooltip_nl_fr, FlagIcon_NL_FR)
        commandGroup2.AddCommand(ButtonText_nl_de, ActionName_nl_de, ButtonDescription_nl_de, ButtonTooltip_nl_de, FlagIcon_NL_DE)

        commandGroup2.AddCommand(ButtonText_de_en, ActionName_de_en, ButtonDescription_de_en, ButtonTooltip_de_en, FlagIcon_DE_EN)
        commandGroup2.AddCommand(ButtonText_de_fr, ActionName_de_fr, ButtonDescription_de_fr, ButtonTooltip_de_fr, FlagIcon_DE_FR)
        commandGroup2.AddCommand(ButtonText_de_it, ActionName_de_it, ButtonDescription_de_it, ButtonTooltip_de_it, FlagIcon_DE_IT)
        commandGroup2.AddCommand(ButtonText_de_pl, ActionName_de_pl, ButtonDescription_de_pl, ButtonTooltip_de_pl, FlagIcon_DE_PL)
        commandGroup2.AddCommand(ButtonText_de_cz, ActionName_de_cz, ButtonDescription_de_cz, ButtonTooltip_de_cz, FlagIcon_DE_CZ)
        commandGroup2.AddCommand(ButtonText_de_hu, ActionName_de_hu, ButtonDescription_de_hu, ButtonTooltip_de_hu, FlagIcon_DE_HU)
        commandGroup2.AddCommand(ButtonText_de_ro, ActionName_de_ro, ButtonDescription_de_ro, ButtonTooltip_de_ro, FlagIcon_DE_RO)
        commandGroup2.AddCommand(ButtonText_de_ru, ActionName_de_ru, ButtonDescription_de_ru, ButtonTooltip_de_ru, FlagIcon_DE_RU)
        commandGroup2.AddCommand(ButtonText_de_se, ActionName_de_se, ButtonDescription_de_se, ButtonTooltip_de_se, FlagIcon_DE_SE)
        commandGroup2.AddCommand(ButtonText_de_da, ActionName_de_da, ButtonDescription_de_da, ButtonTooltip_de_da, FlagIcon_DE_DA)

        commandGroup2.AddCommand(ButtonText_fr_en, ActionName_fr_en, ButtonDescription_fr_en, ButtonTooltip_fr_en, FlagIcon_FR_EN)
        commandGroup2.AddCommand(ButtonText_fr_de, ActionName_fr_de, ButtonDescription_fr_de, ButtonTooltip_fr_de, FlagIcon_FR_DE)
        commandGroup2.AddCommand(ButtonText_fr_nl, ActionName_fr_nl, ButtonDescription_fr_nl, ButtonTooltip_fr_nl, FlagIcon_FR_NL)
        commandGroup2.AddCommand(ButtonText_fr_it, ActionName_fr_it, ButtonDescription_fr_it, ButtonTooltip_fr_it, FlagIcon_FR_IT)
        commandGroup2.AddCommand(ButtonText_fr_es, ActionName_fr_es, ButtonDescription_fr_es, ButtonTooltip_fr_es, FlagIcon_FR_ES)

        commandGroup2.AddCommand(ButtonText_es_en, ActionName_es_en, ButtonDescription_es_en, ButtonTooltip_es_en, FlagIcon_ES_EN)
        commandGroup2.AddCommand(ButtonText_es_fr, ActionName_es_fr, ButtonDescription_es_fr, ButtonTooltip_es_fr, FlagIcon_ES_FR)
        commandGroup2.AddCommand(ButtonText_es_pt, ActionName_es_pt, ButtonDescription_es_pt, ButtonTooltip_es_pt, FlagIcon_ES_PT)
        commandGroup2.AddCommand(ButtonText_es_it, ActionName_es_it, ButtonDescription_es_it, ButtonTooltip_es_it, FlagIcon_ES_IT)

        commandGroup2.AddCommand(ButtonText_it_en, ActionName_it_en, ButtonDescription_it_en, ButtonTooltip_it_en, FlagIcon_IT_EN)
        commandGroup2.AddCommand(ButtonText_it_de, ActionName_it_de, ButtonDescription_it_de, ButtonTooltip_it_de, FlagIcon_IT_DE)
        commandGroup2.AddCommand(ButtonText_it_fr, ActionName_it_fr, ButtonDescription_it_fr, ButtonTooltip_it_fr, FlagIcon_IT_FR)
        commandGroup2.AddCommand(ButtonText_it_es, ActionName_it_es, ButtonDescription_it_es, ButtonTooltip_it_es, FlagIcon_IT_ES)

        commandGroup2.AddCommand(ButtonText_pl_en, ActionName_pl_en, ButtonDescription_pl_en, ButtonTooltip_pl_en, FlagIcon_PL_EN)
        commandGroup2.AddCommand(ButtonText_pl_de, ActionName_pl_de, ButtonDescription_pl_de, ButtonTooltip_pl_de, FlagIcon_PL_DE)
        commandGroup2.AddCommand(ButtonText_pl_cz, ActionName_pl_cz, ButtonDescription_pl_cz, ButtonTooltip_pl_cz, FlagIcon_PL_CZ)

        commandGroup2.AddCommand(ButtonText_cz_en, ActionName_cz_en, ButtonDescription_cz_en, ButtonTooltip_cz_en, FlagIcon_CZ_EN)
        commandGroup2.AddCommand(ButtonText_cz_de, ActionName_cz_de, ButtonDescription_cz_de, ButtonTooltip_cz_de, FlagIcon_CZ_DE)
        commandGroup2.AddCommand(ButtonText_cz_pl, ActionName_cz_pl, ButtonDescription_cz_pl, ButtonTooltip_cz_pl, FlagIcon_CZ_PL)
        commandGroup2.AddCommand(ButtonText_cz_hu, ActionName_cz_hu, ButtonDescription_cz_hu, ButtonTooltip_cz_hu, FlagIcon_CZ_HU)

        commandGroup2.AddCommand(ButtonText_hu_en, ActionName_hu_en, ButtonDescription_hu_en, ButtonTooltip_hu_en, FlagIcon_HU_EN)
        commandGroup2.AddCommand(ButtonText_hu_de, ActionName_hu_de, ButtonDescription_hu_de, ButtonTooltip_hu_de, FlagIcon_HU_DE)
        commandGroup2.AddCommand(ButtonText_hu_cz, ActionName_hu_cz, ButtonDescription_hu_cz, ButtonTooltip_hu_cz, FlagIcon_HU_CZ)
        commandGroup2.AddCommand(ButtonText_hu_ro, ActionName_hu_ro, ButtonDescription_hu_ro, ButtonTooltip_hu_ro, FlagIcon_HU_RO)

        commandGroup2.AddCommand(ButtonText_se_en, ActionName_se_en, ButtonDescription_se_en, ButtonTooltip_se_en, FlagIcon_SE_EN)
        commandGroup2.AddCommand(ButtonText_se_da, ActionName_se_da, ButtonDescription_se_da, ButtonTooltip_se_da, FlagIcon_SE_DA)
        commandGroup2.AddCommand(ButtonText_se_de, ActionName_se_de, ButtonDescription_se_de, ButtonTooltip_se_de, FlagIcon_SE_DE)

        commandGroup2.AddCommand(ButtonText_da_en, ActionName_da_en, ButtonDescription_da_en, ButtonTooltip_da_en, FlagIcon_DA_EN)
        commandGroup2.AddCommand(ButtonText_da_se, ActionName_da_se, ButtonDescription_da_se, ButtonTooltip_da_se, FlagIcon_DA_SE)
        commandGroup2.AddCommand(ButtonText_da_de, ActionName_da_de, ButtonDescription_da_de, ButtonTooltip_da_de, FlagIcon_DA_DE)

        commandGroup2.AddCommand(ButtonText_pt_en, ActionName_pt_en, ButtonDescription_pt_en, ButtonTooltip_pt_en, FlagIcon_PT_EN)
        commandGroup2.AddCommand(ButtonText_pt_es, ActionName_pt_es, ButtonDescription_pt_es, ButtonTooltip_pt_es, FlagIcon_PT_ES)
        commandGroup2.AddCommand(ButtonText_br_en, ActionName_br_en, ButtonDescription_br_en, ButtonTooltip_br_en, FlagIcon_BR_EN)

        commandGroup2.AddCommand(ButtonText_ro_en, ActionName_ro_en, ButtonDescription_ro_en, ButtonTooltip_ro_en, FlagIcon_RO_EN)
        commandGroup2.AddCommand(ButtonText_ro_de, ActionName_ro_de, ButtonDescription_ro_de, ButtonTooltip_ro_de, FlagIcon_RO_DE)
        commandGroup2.AddCommand(ButtonText_ro_hu, ActionName_ro_hu, ButtonDescription_ro_hu, ButtonTooltip_ro_hu, FlagIcon_RO_HU)

        commandGroup2.AddCommand(ButtonText_tr_en, ActionName_tr_en, ButtonDescription_tr_en, ButtonTooltip_tr_en, FlagIcon_TR_EN)
        commandGroup2.AddCommand(ButtonText_ru_en, ActionName_ru_en, ButtonDescription_ru_en, ButtonTooltip_ru_en, FlagIcon_RU_EN)

        commandGroup2.AddCommand(ButtonText_zh_en, ActionName_zh_en, ButtonDescription_zh_en, ButtonTooltip_zh_en, FlagIcon_ZH_EN)
        commandGroup2.AddCommand(ButtonText_ja_en, ActionName_ja_en, ButtonDescription_ja_en, ButtonTooltip_ja_en, FlagIcon_JA_EN)
        commandGroup2.AddCommand(ButtonText_ko_en, ActionName_ko_en, ButtonDescription_ko_en, ButtonTooltip_ko_en, FlagIcon_KO_EN)
        commandGroup2.AddCommand(ButtonText_zh_ja, ActionName_zh_ja, ButtonDescription_zh_ja, ButtonTooltip_zh_ja, FlagIcon_ZH_JA)
        commandGroup2.AddCommand(ButtonText_zh_ko, ActionName_zh_ko, ButtonDescription_zh_ko, ButtonTooltip_zh_ko, FlagIcon_ZH_KO)
        commandGroup2.AddCommand(ButtonText_ja_ko, ActionName_ja_ko, ButtonDescription_ja_ko, ButtonTooltip_ja_ko, FlagIcon_JA_KO)

    End Sub

    ''' <summary>
    ''' Removes the command group from the ribbon when the script
    ''' is unregistered in EPLAN.
    ''' </summary>
    <Eplan.EplApi.Scripting.DeclareUnregister>
    Public Sub UnregisterRibbon()

        Dim eRibbon As New Eplan.EplApi.Gui.RibbonBar
        Dim eTab As Eplan.EplApi.Gui.RibbonTab

        For Each eTab In eRibbon.Tabs
            If eTab.name = TabName1 Or eTab.name = TabName2 Then
                eTab.Remove()
            End If
        Next
    End Sub

#End Region

#Region "Ribbon button actions"

    ' ---------------------------------------------------------------------
    'single language actions
    ' ---------------------------------------------------------------------

    <Eplan.EplApi.Scripting.DeclareAction(ActionName_de_DE)>
Public Sub Action_de_DE()

    'set the project language to German
    ACC.AddParameter("DISPLAY", "de_DE")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_en_US)>
Public Sub Action_en_US()

    'set the project language to English
    ACC.AddParameter("DISPLAY", "en_US")
    ACC.AddParameter("VARIABLE", "en_US")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_es_ES)>
Public Sub Action_es_ES()

    'set the project language to Spanish
    ACC.AddParameter("DISPLAY", "es_ES")
    ACC.AddParameter("VARIABLE", "es_ES")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_fr_FR)>
Public Sub Action_fr_FR()

    'set the project language to French
    ACC.AddParameter("DISPLAY", "fr_FR")
    ACC.AddParameter("VARIABLE", "fr_FR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_nl_NL)>
Public Sub Action_nl_NL()

    'set the project language to Dutch
    ACC.AddParameter("DISPLAY", "nl_NL")
    ACC.AddParameter("VARIABLE", "nl_NL")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_sv_SE)>
Public Sub Action_sv_SE()

    'set the project language to Swedish
    ACC.AddParameter("DISPLAY", "sv_SE")
    ACC.AddParameter("VARIABLE", "sv_SE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_da_DK)>
Public Sub Action_da_DK()

    'set the project language to Danish
    ACC.AddParameter("DISPLAY", "da_DK")
    ACC.AddParameter("VARIABLE", "da_DK")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ru_RU)>
Public Sub Action_ru_RU()

    'set the project language to Russian
    ACC.AddParameter("DISPLAY", "ru_RU")
    ACC.AddParameter("VARIABLE", "ru_RU")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_zh_CN)>
Public Sub Action_zh_CN()

    'set the project language to Chinese
    ACC.AddParameter("DISPLAY", "zh_CN")
    ACC.AddParameter("VARIABLE", "zh_CN")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_pl_PL)>
Public Sub Action_pl_PL()

    'set the project language to Polish
    ACC.AddParameter("DISPLAY", "pl_PL")
    ACC.AddParameter("VARIABLE", "pl_PL")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_pt_BR)>
Public Sub Action_pt_BR()

    'set the project language to Portuguese Brazil
    ACC.AddParameter("DISPLAY", "pt_BR")
    ACC.AddParameter("VARIABLE", "pt_BR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_cs_CZ)>
Public Sub Action_cs_CZ()

    'set the project language to Czech
    ACC.AddParameter("DISPLAY", "cs_CZ")
    ACC.AddParameter("VARIABLE", "cs_CZ")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_it_IT)>
Public Sub Action_it_IT()

    'set the project language to Italian
    ACC.AddParameter("DISPLAY", "it_IT")
    ACC.AddParameter("VARIABLE", "it_IT")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_hu_HU)>
Public Sub Action_hu_HU()

    'set the project language to Hungarian
    ACC.AddParameter("DISPLAY", "hu_HU")
    ACC.AddParameter("VARIABLE", "hu_HU")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_pt_PT)>
Public Sub Action_pt_PT()

    'set the project language to Portuguese Portugal
    ACC.AddParameter("DISPLAY", "pt_PT")
    ACC.AddParameter("VARIABLE", "pt_PT")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ko_KR)>
Public Sub Action_ko_KR()

    'set the project language to Korean
    ACC.AddParameter("DISPLAY", "ko_KR")
    ACC.AddParameter("VARIABLE", "ko_KR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ja_JP)>
Public Sub Action_ja_JP()

    'set the project language to Japanese
    ACC.AddParameter("DISPLAY", "ja_JP")
    ACC.AddParameter("VARIABLE", "ja_JP")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_tr_TR)>
Public Sub Action_tr_TR()

    'set the project language to Turkish
    ACC.AddParameter("DISPLAY", "tr_TR")
    ACC.AddParameter("VARIABLE", "tr_TR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ro_RO)>
Public Sub Action_ro_RO()

    'set the project language to Romanian
    ACC.AddParameter("DISPLAY", "ro_RO")
    ACC.AddParameter("VARIABLE", "ro_RO")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

' ---------------------------------------------------------------------
'dual language actions
' ---------------------------------------------------------------------


<Eplan.EplApi.Scripting.DeclareAction(ActionName_nl_en)>
Public Sub Action_nl_en()
    'set the project language to Dutch and English
    ACC.AddParameter("DISPLAY", "nl_NL;en_US")
    ACC.AddParameter("VARIABLE", "nl_NL")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_nl_fr)>
Public Sub Action_nl_fr()
    'set the project language to Dutch and French
    ACC.AddParameter("DISPLAY", "nl_NL;fr_FR")
    ACC.AddParameter("VARIABLE", "nl_NL")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_nl_de)>
Public Sub Action_nl_de()
    'set the project language to Dutch and German
    ACC.AddParameter("DISPLAY", "nl_NL;de_DE")
    ACC.AddParameter("VARIABLE", "nl_NL")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_en)>
Public Sub Action_de_en()
    'set the project language to German and English
    ACC.AddParameter("DISPLAY", "de_DE;en_US")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_fr)>
Public Sub Action_de_fr()
    'set the project language to German and French
    ACC.AddParameter("DISPLAY", "de_DE;fr_FR")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_it)>
Public Sub Action_de_it()
    'set the project language to German and Italian
    ACC.AddParameter("DISPLAY", "de_DE;it_IT")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_pl)>
Public Sub Action_de_pl()
    'set the project language to German and Polish
    ACC.AddParameter("DISPLAY", "de_DE;pl_PL")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_cz)>
Public Sub Action_de_cz()
    'set the project language to German and Czech
    ACC.AddParameter("DISPLAY", "de_DE;cs_CZ")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_hu)>
Public Sub Action_de_hu()
    'set the project language to German and Hungarian
    ACC.AddParameter("DISPLAY", "de_DE;hu_HU")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_ro)>
Public Sub Action_de_ro()
    'set the project language to German and Romanian
    ACC.AddParameter("DISPLAY", "de_DE;ro_RO")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_ru)>
Public Sub Action_de_ru()
    'set the project language to German and Russian
    ACC.AddParameter("DISPLAY", "de_DE;ru_RU")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_se)>
Public Sub Action_de_se()
    'set the project language to German and Swedish
    ACC.AddParameter("DISPLAY", "de_DE;sv_SE")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_de_da)>
Public Sub Action_de_da()
    'set the project language to German and Danish
    ACC.AddParameter("DISPLAY", "de_DE;da_DK")
    ACC.AddParameter("VARIABLE", "de_DE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_fr_en)>
Public Sub Action_fr_en()
    'set the project language to French and English
    ACC.AddParameter("DISPLAY", "fr_FR;en_US")
    ACC.AddParameter("VARIABLE", "fr_FR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_fr_de)>
Public Sub Action_fr_de()
    'set the project language to French and German
    ACC.AddParameter("DISPLAY", "fr_FR;de_DE")
    ACC.AddParameter("VARIABLE", "fr_FR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_fr_nl)>
Public Sub Action_fr_nl()
    'set the project language to French and Dutch
    ACC.AddParameter("DISPLAY", "fr_FR;nl_NL")
    ACC.AddParameter("VARIABLE", "fr_FR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_fr_it)>
Public Sub Action_fr_it()
    'set the project language to French and Italian
    ACC.AddParameter("DISPLAY", "fr_FR;it_IT")
    ACC.AddParameter("VARIABLE", "fr_FR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_fr_es)>
Public Sub Action_fr_es()
    'set the project language to French and Spanish
    ACC.AddParameter("DISPLAY", "fr_FR;es_ES")
    ACC.AddParameter("VARIABLE", "fr_FR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_es_en)>
Public Sub Action_es_en()
    'set the project language to Spanish and English
    ACC.AddParameter("DISPLAY", "es_ES;en_US")
    ACC.AddParameter("VARIABLE", "es_ES")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_es_fr)>
Public Sub Action_es_fr()
    'set the project language to Spanish and French
    ACC.AddParameter("DISPLAY", "es_ES;fr_FR")
    ACC.AddParameter("VARIABLE", "es_ES")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_es_pt)>
Public Sub Action_es_pt()
    'set the project language to Spanish and Portuguese
    ACC.AddParameter("DISPLAY", "es_ES;pt_PT")
    ACC.AddParameter("VARIABLE", "es_ES")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_es_it)>
Public Sub Action_es_it()
    'set the project language to Spanish and Italian
    ACC.AddParameter("DISPLAY", "es_ES;it_IT")
    ACC.AddParameter("VARIABLE", "es_ES")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_it_en)>
Public Sub Action_it_en()
    'set the project language to Italian and English
    ACC.AddParameter("DISPLAY", "it_IT;en_US")
    ACC.AddParameter("VARIABLE", "it_IT")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_it_de)>
Public Sub Action_it_de()
    'set the project language to Italian and German
    ACC.AddParameter("DISPLAY", "it_IT;de_DE")
    ACC.AddParameter("VARIABLE", "it_IT")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_it_fr)>
Public Sub Action_it_fr()
    'set the project language to Italian and French
    ACC.AddParameter("DISPLAY", "it_IT;fr_FR")
    ACC.AddParameter("VARIABLE", "it_IT")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_it_es)>
Public Sub Action_it_es()
    'set the project language to Italian and Spanish
    ACC.AddParameter("DISPLAY", "it_IT;es_ES")
    ACC.AddParameter("VARIABLE", "it_IT")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_pl_en)>
Public Sub Action_pl_en()
    'set the project language to Polish and English
    ACC.AddParameter("DISPLAY", "pl_PL;en_US")
    ACC.AddParameter("VARIABLE", "pl_PL")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_pl_de)>
Public Sub Action_pl_de()
    'set the project language to Polish and German
    ACC.AddParameter("DISPLAY", "pl_PL;de_DE")
    ACC.AddParameter("VARIABLE", "pl_PL")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_pl_cz)>
Public Sub Action_pl_cz()
    'set the project language to Polish and Czech
    ACC.AddParameter("DISPLAY", "pl_PL;cs_CZ")
    ACC.AddParameter("VARIABLE", "pl_PL")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_cz_en)>
Public Sub Action_cz_en()
    'set the project language to Czech and English
    ACC.AddParameter("DISPLAY", "cs_CZ;en_US")
    ACC.AddParameter("VARIABLE", "cs_CZ")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_cz_de)>
Public Sub Action_cz_de()
    'set the project language to Czech and German
    ACC.AddParameter("DISPLAY", "cs_CZ;de_DE")
    ACC.AddParameter("VARIABLE", "cs_CZ")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_cz_pl)>
Public Sub Action_cz_pl()
    'set the project language to Czech and Polish
    ACC.AddParameter("DISPLAY", "cs_CZ;pl_PL")
    ACC.AddParameter("VARIABLE", "cs_CZ")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_cz_hu)>
Public Sub Action_cz_hu()
    'set the project language to Czech and Hungarian
    ACC.AddParameter("DISPLAY", "cs_CZ;hu_HU")
    ACC.AddParameter("VARIABLE", "cs_CZ")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_hu_en)>
Public Sub Action_hu_en()
    'set the project language to Hungarian and English
    ACC.AddParameter("DISPLAY", "hu_HU;en_US")
    ACC.AddParameter("VARIABLE", "hu_HU")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_hu_de)>
Public Sub Action_hu_de()
    'set the project language to Hungarian and German
    ACC.AddParameter("DISPLAY", "hu_HU;de_DE")
    ACC.AddParameter("VARIABLE", "hu_HU")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_hu_cz)>
Public Sub Action_hu_cz()
    'set the project language to Hungarian and Czech
    ACC.AddParameter("DISPLAY", "hu_HU;cs_CZ")
    ACC.AddParameter("VARIABLE", "hu_HU")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_hu_ro)>
Public Sub Action_hu_ro()
    'set the project language to Hungarian and Romanian
    ACC.AddParameter("DISPLAY", "hu_HU;ro_RO")
    ACC.AddParameter("VARIABLE", "hu_HU")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_se_en)>
Public Sub Action_se_en()
    'set the project language to Swedish and English
    ACC.AddParameter("DISPLAY", "sv_SE;en_US")
    ACC.AddParameter("VARIABLE", "sv_SE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_se_da)>
Public Sub Action_se_da()
    'set the project language to Swedish and Danish
    ACC.AddParameter("DISPLAY", "sv_SE;da_DK")
    ACC.AddParameter("VARIABLE", "sv_SE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_se_de)>
Public Sub Action_se_de()
    'set the project language to Swedish and German
    ACC.AddParameter("DISPLAY", "sv_SE;de_DE")
    ACC.AddParameter("VARIABLE", "sv_SE")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_da_en)>
Public Sub Action_da_en()
    'set the project language to Danish and English
    ACC.AddParameter("DISPLAY", "da_DK;en_US")
    ACC.AddParameter("VARIABLE", "da_DK")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_da_se)>
Public Sub Action_da_se()
    'set the project language to Danish and Swedish
    ACC.AddParameter("DISPLAY", "da_DK;sv_SE")
    ACC.AddParameter("VARIABLE", "da_DK")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_da_de)>
Public Sub Action_da_de()
    'set the project language to Danish and German
    ACC.AddParameter("DISPLAY", "da_DK;de_DE")
    ACC.AddParameter("VARIABLE", "da_DK")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_pt_en)>
Public Sub Action_pt_en()
    'set the project language to Portuguese and English
    ACC.AddParameter("DISPLAY", "pt_PT;en_US")
    ACC.AddParameter("VARIABLE", "pt_PT")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_pt_es)>
Public Sub Action_pt_es()
    'set the project language to Portuguese and Spanish
    ACC.AddParameter("DISPLAY", "pt_PT;es_ES")
    ACC.AddParameter("VARIABLE", "pt_PT")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_br_en)>
Public Sub Action_br_en()
    'set the project language to Portuguese Brazil and English
    ACC.AddParameter("DISPLAY", "pt_BR;en_US")
    ACC.AddParameter("VARIABLE", "pt_BR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ro_en)>
Public Sub Action_ro_en()
    'set the project language to Romanian and English
    ACC.AddParameter("DISPLAY", "ro_RO;en_US")
    ACC.AddParameter("VARIABLE", "ro_RO")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ro_de)>
Public Sub Action_ro_de()
    'set the project language to Romanian and German
    ACC.AddParameter("DISPLAY", "ro_RO;de_DE")
    ACC.AddParameter("VARIABLE", "ro_RO")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ro_hu)>
Public Sub Action_ro_hu()
    'set the project language to Romanian and Hungarian
    ACC.AddParameter("DISPLAY", "ro_RO;hu_HU")
    ACC.AddParameter("VARIABLE", "ro_RO")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_tr_en)>
Public Sub Action_tr_en()
    'set the project language to Turkish and English
    ACC.AddParameter("DISPLAY", "tr_TR;en_US")
    ACC.AddParameter("VARIABLE", "tr_TR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ru_en)>
Public Sub Action_ru_en()
    'set the project language to Russian and English
    ACC.AddParameter("DISPLAY", "ru_RU;en_US")
    ACC.AddParameter("VARIABLE", "ru_RU")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_zh_en)>
Public Sub Action_zh_en()
    'set the project language to Chinese and English
    ACC.AddParameter("DISPLAY", "zh_CN;en_US")
    ACC.AddParameter("VARIABLE", "zh_CN")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ja_en)>
Public Sub Action_ja_en()
    'set the project language to Japanese and English
    ACC.AddParameter("DISPLAY", "ja_JP;en_US")
    ACC.AddParameter("VARIABLE", "ja_JP")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ko_en)>
Public Sub Action_ko_en()
    'set the project language to Korean and English
    ACC.AddParameter("DISPLAY", "ko_KR;en_US")
    ACC.AddParameter("VARIABLE", "ko_KR")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_zh_ja)>
Public Sub Action_zh_ja()
    'set the project language to Chinese and Japanese
    ACC.AddParameter("DISPLAY", "zh_CN;ja_JP")
    ACC.AddParameter("VARIABLE", "zh_CN")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_zh_ko)>
Public Sub Action_zh_ko()
    'set the project language to Chinese and Korean
    ACC.AddParameter("DISPLAY", "zh_CN;ko_KR")
    ACC.AddParameter("VARIABLE", "zh_CN")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub

<Eplan.EplApi.Scripting.DeclareAction(ActionName_ja_ko)>
Public Sub Action_ja_ko()
    'set the project language to Japanese and Korean
    ACC.AddParameter("DISPLAY", "ja_JP;ko_KR")
    ACC.AddParameter("VARIABLE", "ja_JP")
    CLI.Execute("SetProjectLanguage", ACC) 'Hoofdlettergevoelig, 
End Sub


#End Region

End Class

#Region "Icon Reference"

' =========================================================================
' Available EPLAN Ribbon Icons
' =========================================================================
'
' Circles
' --------
' Circle_0 .. Circle_9
'
' Octagons
' --------
' Octagon_0 .. Octagon_9
'
' Diamonds
' --------
' Diamond_0 .. Diamond_9
'
' Hexagons
' --------
' Hexagon_0 .. Hexagon_9
'
' Rectangles (Numbers)
' --------------------
' Rectangle_0 .. Rectangle_9
'
' Rectangles (Letters)
' --------------------
' Rectangle_A .. Rectangle_Z
'
' EPLAN / System
' --------------
' Eplan
' SystemConfig
' Application
' TaskList
' SymbolRepresentationType
'
' Electrical
' ----------
' TransformerSingle
' WireSingle
' PartCable
' PartPLC
' PartTerminal
' PartPlug
' Relay
' Generator
' Converter
' Capacitor
' Inductor
' Resistor
' Motor
' Pump
' PowerUnit
' SafetyDevice
' SignalingDevice
' Switchgear
' SwitchSingle
' ControlTerminal
' MeasuringDevice
' MeasuringConnection
'
' Mechanical / Fluid
' ------------------
' FluidDrive
' FluidFilter
' FluidMeasuringDevice
' FluidSensor
' FluidValve
' FluidPowerArea
' PartCylinder
' PartHeating
' PartRectifier
' Pump
' ValveSingle
'
' Process / P&I
' -------------
' Piping
' PiPiping
' PiVessel
' PiStorage
' PiMixer
' PiStirrer
' PiCooler
' PiFilter
' PiSeparator
' PiConveyor
' PiCompressor
' PiChimney
' PiState
' PiMedium
'
' PPE
' ---
' PpePlant
' PpePlantPart
' PpePlantSection
' PpeMachine
' PpeFunction
' PpeSensor
' PpeSoftwareFunction
' PpeConsumer
' PpeBlockOfBuildings
' PpeWorkBuildings
'
' Flags
' -----
' Flag_AU
' Flag_BR
' Flag_CH
' Flag_CZ
' Flag_DA
' Flag_DE
' Flag_EN
' Flag_ES
' Flag_FR
' Flag_HU
' Flag_IT
' Flag_JA
' Flag_KO
' Flag_NL
' Flag_PL
' Flag_PT
' Flag_RO
' Flag_RU
' Flag_SE
' Flag_TR
' Flag_US
' Flag_ZH
'
' Combined Flags
' --------------
' Flag_DE_BR
' Flag_DE_CZ
' Flag_DE_DA
' Flag_DE_EN
' Flag_DE_ES
' Flag_DE_FR
' Flag_DE_GB
' Flag_DE_HU
' Flag_DE_IT
' Flag_DE_KO
' Flag_DE_NL
' Flag_DE_PL
' Flag_DE_PT
' Flag_DE_RU
' Flag_DE_SE
' Flag_DE_US
' Flag_DE_ZH
' Flag_EN_DE
' Flag_IT_DE
' Flag_IT_US
' Flag_US_DE
' Flag_US_IT
'
' Recommended examples
' --------------------
' TaskList
' Application
' SystemConfig
' Eplan
' PartPLC
' PartTerminal
' Rectangle_A
' Rectangle_K
' Circle_1
' Hexagon_7
' Flag_NL
'
' =========================================================================

#End Region