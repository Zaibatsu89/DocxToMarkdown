# DocxToMarkdown

DocxToMarkdown is een C#-applicatie waarmee je eenvoudig Microsoft Word-documenten (.docx) kunt converteren naar Markdown-formaat. Deze tool is ideaal voor ontwikkelaars, schrijvers en documentatiebeheerders die hun Word-bestanden snel willen omzetten voor gebruik in bijvoorbeeld GitHub, wikis of andere markdown-omgevingen.

## Features

- Converteert .docx-bestanden naar goed gestructureerde Markdown-bestanden
- Ondersteunt tekst, koppen, lijsten, tabellen en afbeeldingen
- Batchverwerking voor meerdere bestanden tegelijk
- Eenvoudige command-line interface voor automatisering
- Snel, lichtgewicht en gemakkelijk te gebruiken

## Installatie

1. Clone deze repository:
   ```bash
   git clone https://github.com/Zaibatsu89/DocxToMarkdown.git
   ```
2. Open het project in Visual Studio of een andere C# IDE.
3. Herstel de benodigde NuGet-packages.
4. Bouw het project.

## Gebruik

1. Zorg dat je een .docx-bestand hebt dat je wilt converteren.
2. Start de applicatie via de command line:

   ```bash
   DocxToMarkdown.exe -i "<pad/naar/input.docx>" -o "<pad/naar/output.md>"
   ```

**Voorbeeld:**
```bash
DocxToMarkdown.exe -i "C:\Documenten\voorbeeld.docx" -o "C:\Documenten\voorbeeld.md"
```

### Opties

- `-i`, `--input` : Pad naar het .docx-bestand (verplicht)
- `-o`, `--output` : Pad voor het gegenereerde .md-bestand (verplicht)
- `--batch` : Map met meerdere .docx-bestanden voor batchverwerking (optioneel)

## Voorbeeld output

```markdown
# Titel uit Word-document

Dit is een paragraaf uit het document.

## Tussenkopje

- Eerste punt
- Tweede punt
```

## Bijdragen

Bijdragen zijn welkom! Open een issue voor bug reports of feature requests, of maak een pull request met verbeteringen.

## Licentie

Dit project is gelicenseerd onder de MIT-licentie.

## Contact

Voor vragen of feedback, maak gerust een issue aan op GitHub of neem contact op via [Zaibatsu89](https://github.com/Zaibatsu89).
