# Spreadsheet programming - Excel VBA and Google Sheets GAS comparisons

## Read on if:

- You are interested in spreadsheet programming.
- You are an Excel VBA programmer interested in learning how to program Google Sheets GAS (Google App Script).
- You already know  Google Sheets GAS and would like to learn Excel VBA.
- You are an Excel VBA programmer who would like to learn more about JavaScript


GAS is essentially a modern version JavaScript so it is very different from VBA both in terms of syntax and semantics. VBA is rather like a living fossil in that it has hardly changed since the 1990s. In describing VBA as a living fossil, my intention is not to belittle it. Despite its shortcomings, it is still a very effective tool for its intended purpose of extending and customising Excel functionality. GAS and Sheets are both much more recent than Excel and its hosted VBA and Google upgraded its JavaScript support recently so most modern JavaScript features are available.

The purpose of the comparison here is not to persuade you that one language is better than the other for the purpose at hand but rather to show each language can be used to achieve the same objective. JavaScript is clearly a very important language given its dominance of the browser so it is certainly worth learning.

__A note on terminology__: I use the terms GAS and JavaScript interchangeably. When I use _JavaScript_ I am referring to a JavaScript feature that applies to the language generally while when I write of _GAS_, I am referring to a specific feature relevant to the Google Sheets implementation only. For example arrays and loops apply to JavaScript anywhere whether in GAS, the browser or whatever while a sheets _range_ of cells is specific to the GAS implementation of the language.

## Markdown generation

The first task I have chosen develop in parallel in both languages is a custom function (aka user-defined function) that can convert a range input of values into a Markdown table. 

The explanatory notes comparing VBA and GAS explore:

- Custom functions
- Variable declarations
- Text handing
- Arrays

The full discussion: [Markdown table generation using a user-defined function](https://github.com/Rotifer/VBA_GAS/blob/main/Markdown_Generation/notes_on_generation_of_md_table.md)