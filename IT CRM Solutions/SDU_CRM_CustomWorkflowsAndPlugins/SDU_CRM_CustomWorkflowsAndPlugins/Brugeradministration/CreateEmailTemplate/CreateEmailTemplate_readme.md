# CreateEmailTemplateeeeee123
## Formål
Koden har til formål at generere email inhold til brug ved bestillingen af rettigheder. I nogle få tilfælde
bliver emailen afsendt indefra koden (ØSS og HR), da det er nødvendigt at CC et dynamisk antal af institutledere.
I alle andre tilfælde bliver koden kaldt fra et workflow i CRM, hvorefter outputs fra koden bruges til at udfylde en email.

## Funktionalitet
Koden kaldes fra workflows i CRM, hvori en række parametre bruges. Der bruges bl.a. to logiske parametre *ExistingRights*
og *RemoveRightsCompletely*. Disse bruges til at udlede om der er tale om en fjernelse eller ændring/tilføjelse af rettigheder.
Dette udledes dermed ikke eksplicit i koden, men gøres i workflowet, hvori de enkelte steps medsender forskellige TRUE/FALSE værdier
i de to parametre. 

Baseret på ovenstående så igangssættes ét af tre metoder, som genererer indholdet til emailen. Udover det anvendes der en metode
til at fremsøge relaterede records, som dynamisk skal indsættes som tekst i emailen. Man kan have x-anta af ØSS, HR, Qlikview samt
Acadre rettigheder.

### Metoder
#### Email generering
Der anvendes nedenstående tre metoder til at generere email indhold. De returnerer alle en Tuple<string, string> med en emailbody
og et subjekt.

Ved de to første (nye eller ændringer) der medsendes en *DataCollection\<Entity>* som indeholder de relaterede records som skal indsættes
i emailen.

- buildEmailBody_NoExistingRights
- buildEmailBody_ChangesInRights
- buildEmailBody_RemoveAllRights

#### Fremsøgning af relaterede records
- QueryForRelatedRecords
- getRelatedRecords

Førstnævnte metode fremsøger de relaterede records. Der medsendes heri bl.a. et string parameter, som fastslår hvilken
entitet man ønsker at lede efter. Recordsne findes dermed altid ift. den igangsættende record og records i den entitet
man angiver.

*getRelatedRecords* modtager collectionen af records og anvender en StringBuilder til at generere dynamisk indhold til emailen.

#### Afsendelse af email
- SendEmail

Metoden anvendes udelukkende ved HR og QlikView, da det her er påkrævet, at den enkelte institutleder er CC på den email der afsendes.
En person kan have ØSS/HR rettigheder på X antal af omkostningssteder, hvor institutlederen for hver skal CC på den samme email.

## Fleksibilitet og fremtidige ændringer
Metoderne er anvendes heri er fleksibile ift. de parametre der afsendes, og nye rettighedsområder er derfor ikke vanskelige at indsætte.
De seks rettighedsområder som i dag (28-01-2020) er alle i dag lavet ved Late Bound relationer til felter i bestillingen. 
Hvis der i fremtiden kommer nye felter/forsvinder felter, så er det nødvendigt at ændre i det objekt, som medsendes i email metoderne (som afgør hvilke felter der skal fremgå af emailen).

Det kunne være elegant hvis der blev etableret en mere intuitiv måde at "fodre" emailgenerering på, så det er ikke var nødvendigt
manuelt at tilføje key-value par til de objekter, som bestemmer indholdet i emailsne. 

## Changelog 
28-01-2020: Log oprettet.