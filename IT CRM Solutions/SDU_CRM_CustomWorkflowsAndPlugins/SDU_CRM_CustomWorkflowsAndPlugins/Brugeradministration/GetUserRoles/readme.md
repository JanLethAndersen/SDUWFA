# GetUserRole + GetUserRole_WFA
## Formål
Koden har til formål at give et overblik over en brugers samtlige rettigheder i CRM.
Disse resultater returneres så de er læsbare, og de evalueres således at det udledes om man har brug
for en *Sales* eller *Team Member*-Licens. 

Koden returnere ligeledes også den sidste login dato, hvilket muliggør en hurtig og potentiel automatisk
deaktivering af brugere baseret på deres seneste login.

Koden findes både som plugin og som CWA (den med prefikset).

## Funktionalitet
Koden anvender en række *QueryExpressions* til at fremsøge *privileges*. Privilegier tilhører
en hænger sammen med en sikkerhedsrolle. Det er derfor nødvendigt at indhente brugerens
sikkerhedsroller, samt de sikkerhedsroller brugeren har gennem teams. Sikkerhedsrollerne
eksisterer i mange forskellige versioner baseret på den BU brugeren tilhører. En metode sikrer
derfor at det er "master" sikkerhedsrollen der tages udgangspunkt, altså den som tilhører den
øverste Business Unit (SDUPRO). Det er kun her man kan fremsøge privilegierne.

Den seneste login dato indhentes ved at anvende classen *RetrieveRecordChangeHistoryRequest*.
Man kan herved forespørge blive alle audit events på den givne bruger, og dernæst filtrere 
dem.

### Metoder
#### getMasterRole

#### getPrivileges

#### getRolesFromTeam

#### CheckLastLoginDate

## Fleksibilitet og fremtidige ændringer

## Changelog 
14-02-2020: Log oprettet.