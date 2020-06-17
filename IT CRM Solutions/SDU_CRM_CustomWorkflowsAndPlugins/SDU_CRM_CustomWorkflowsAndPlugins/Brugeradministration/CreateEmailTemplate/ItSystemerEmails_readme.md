# ITSystemerEmails
## Formål
Formålet med koden er at fremsøge alle de IT-rettigheder, som en person
har bestilt i CRM, grupperer dem og udsende mails omkring godkendelsen af dem. 

## Funktionalitet
De relaterede records indhentes via metoden 'QueryForRelatedRecords' i CreateEmailTemplate.cs.
De grupperes herefter baseret på det IT-system de tilhører. For at holde
styr på dem oprettes der et *RightsObject* (class) for hver rettighed, som alle puttes
i en liste. Denne liste grupperes baseret på IT-systemet, og for hver gruppe
identificeres det om rettigheden kræver en institutledergodkendelse.
Hvis de går, så afsendes der hertil - hvis ikke, så afsendes der en email til en systemejer.


### Metoder
#### GroupRightsTogether
For hver record oprettes der et element af classen *RightsObject*, som alle proppes
i en liste. Denne liste grupperes baseret på hvilket IT-system hvert objekt tilhører.
Medlemmerne af hver gruppe sendes videre til ConstructEmailBody.

#### ConstructEmailBody
Baseret på de modtagne objekter, så genereres der en emailtekst via en StringBuilder.
Rettighederne grupperes i hvilke der skal tildeles, og hvilke der skal fratages.

#### SendEmail
Metoden sender emailen, dvs. at selve mailudsendelsen håndteres direkte i koden.

## Fleksibilitet og fremtidige ændringer
Koden er fleksibel i den facon at uansat hvilke IT-rettigheder der måtte blive oprettet
i CRM, så vil denne kode altid virke, da hvert IT-system behandles særskilt. 
Det er samtidigt også hele udviklingsgrundlaget, således at man ikke behøves at ændre
i koden hver gang et nyt system kommer til.

Ift. fremtidige ændringer: refactoring...

## Changelog 
02-04-2020: Log oprettet.