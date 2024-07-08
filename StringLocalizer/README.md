# StringLocalizer

Simple strings localizer.

## How to use

1. Import all modules into your project.
2. Add reference to `Microsoft Scripting Runtime` for your project (because we use dictionaries).
3. Use module `LocalizedStringsEN_Sample.cls` as a boilerplate sample for strings matching. You need to create at least one localization for your project as a fallback localization. So it will be used when there's no match localization module for current Corel locale.
For example, we may create English localization as fallback. Rename module `LocalizedStringsEN_Sample` to, let say, `LocalizedStringsEN` for convience, and create "name-string" pairs for inner `Strings` dictionary, in a manner it currently there. Key in that dictionary will be a name for your placeholder, where string substitutes, and value will be that string.
4. The convient way is to declare a global variable for localizer, but you also may use local variable and pass it between modules of your project, if you against globals. Let's name it `LocalizedStrings` for the following examples.
5. Substitute strings in your project with your localizer variable with `Item` method: `LocalizedStrings.Item("PlaceholderName")`. It's a default method, so it can be omitted: `LocalizedStrings("PlaceholderName")`.
6. (optionally) Add more locales. Create new class module and copy content of your fallback module there. Rename it to somewhat convient like `LocalizedStringsRU`, optionally change comment in header to Corel enum for that local and LCID (locale ID), you may also add localizator credits there:
```VBA
'===============================================================================
' cdrRussian (1049) by Ivan Petrov (https://ivanpetrov.ru)
'===============================================================================
```
Replace copied fallback strings with strings for that locale.

7. Initialize localizer in your main routine like that:
```VBA
With StringLocalizer.Builder(cdrEnglishUS, New LocalizedStringsEN)
    .WithLocale 0000, New LocalizedStringsSomeLocale
    .WithLocale 0001, New LocalizedStringsOtherLocale
    'and so on
    Set LocalizedStrings = .Build
End With
```
Where `cdrEnglishUS`, `0000`, `0001` is LCIDs (you may also use Corel enums), `LocalizedStringsEN` is fallback locale, `LocalizedStringsSomeLocale` and `LocalizedStringsOtherLocale` is additional locales.
Minimal initialization with one locale will be

```VBA
Set LocalizedStrings = StringLocalizer.Builder(cdrEnglishUS, New LocalizedStringsEN).Build
```
**Important**: if you use global variable, set it to `Nothing` at the end of your main routine to terminate `StringLocalizer`.

## Initialization example

```VBA
' In main module
Public LocalizedStrings As IStringLocalizer

' In main routine
Sub Main()
    ' call separate method to initialize
    LocalizedStringsInit
    ' ...
    ' and at the end
    Set LocalizedStrings = Nothing
End Sub

' Initialization method
Private Sub LocalizedStringsInit()
    With StringLocalizer.Builder(cdrEnglishUS, New LocalizedStringsEN)
        .WithLocale cdrRussian, New LocalizedStringsRU
        .WithLocale cdrBrazilianPortuguese, New LocalizedStringsBR
        .WithLocale 1036, New LocalizedStringsFR
        Set LocalizedStrings = .Build
    End With
End Sub
```

## Substitution example

```VBA
BtnCancel.Caption = LocalizedStrings("View.BtnCancel")
' ...
VBA.MsgBox LocalizedStrings("Common.ErrNothingSelected")
```

## Dependences

None.
