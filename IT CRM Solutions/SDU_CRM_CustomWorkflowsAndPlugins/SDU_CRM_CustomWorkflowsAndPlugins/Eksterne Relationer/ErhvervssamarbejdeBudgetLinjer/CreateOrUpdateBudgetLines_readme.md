# CreateOrUpdateBudgetLines
## Formål
Formålet med pluginet er at create records i entiteten "Erhvervssamarbejde budgetlinjer".
Disse oprettes allerede for erhvervssamarbejder på et fakultetsprojekt. Denne kode opretter
dem for selve fakultetsprojektet.

## Funktionalitet
Koden finder ud af hvor mange år et projekt løber over, og kigger på hvad 
bevilling og medfinansiering fra insitut er pr. år. Dvs. beløb/antal år. Afhængigt af 
antallet oprettes budget linjer, som indsættes i tilhørende lookup på fak. projektet.
Der kan oprettes op til 10 (ligesom med erhvervssamarbejde).

Hvilket felt der hører til hvilket år er defineret i listen *fieldIndexDefinition* i koden.

### Metoder


## Fleksibilitet og fremtidige ændringer
- antal år fleksibelt

## Changelog 
12-06-2020: Log oprettet.