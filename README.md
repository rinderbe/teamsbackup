# teamsbackup

PowerShell-modul för backup och återläsning av Microsoft Teams Phone- och Direct Routing-konfiguration.

Modulen är tänkt att köras före ett ändringsarbete i en kunds Teams-miljö. Om ändringen går dåligt kan konfigurationen läsas tillbaka, helt eller delvis. Innan något ändras kan man göra en torrsimning som visar exakt vad som skulle göras.

## Innehåll

- [Krav](#krav)
- [Snabbstart](#snabbstart)
- [Arbetsgång vid ändringsarbete](#arbetsgång-vid-ändringsarbete)
- [Backup-TeamsConfig](#backup-teamsconfig)
- [Restore-TeamsConfig](#restore-teamsconfig)
- [Torrsimning med -WhatIf](#torrsimning-med--whatif)
- [Get-TeamsConfigBackup](#get-teamsconfigbackup)
- [Compare-TeamsConfigBackup](#compare-teamsconfigbackup)
- [Vad som ingår i backupen](#vad-som-ingår-i-backupen)
- [Vad som inte ingår](#vad-som-inte-ingår)
- [Begränsningar och saker att tänka på](#begränsningar-och-saker-att-tänka-på)
- [Felsökning](#felsökning)

## Krav

| Krav | Kommentar |
|---|---|
| PowerShell 7 eller senare | Windows PowerShell 5.1 stöds inte. |
| Modulen `MicrosoftTeams` 6.x eller senare | `Install-Module MicrosoftTeams -Scope CurrentUser` |
| Rollen Teams Administrator | Krävs både för att läsa och för att skriva tillbaka. |

## Snabbstart

```powershell
Import-Module .\TeamsConfigBackup.psm1
Connect-MicrosoftTeams

# Ta backup
Backup-TeamsConfig -Path C:\Kund\Backup -IncludeEffectivePolicies -IncludeUserCallingSettings

# Visa vad som skiljer mot backupen (ändrar ingenting)
Restore-TeamsConfig -Path C:\Kund\Backup\TeamsBackup_Kund_20261018-090000 -WhatIf

# Läs tillbaka allt
Restore-TeamsConfig -Path C:\Kund\Backup\TeamsBackup_Kund_20261018-090000
```

## Arbetsgång vid ändringsarbete

1. Anslut till kundens tenant med `Connect-MicrosoftTeams`.
2. Ta backup med `Backup-TeamsConfig -IncludeEffectivePolicies -IncludeUserCallingSettings`.
3. Kontrollera backupen med `Get-TeamsConfigBackup`. Titta särskilt efter rader märkta `Fel:`.
4. Genomför ändringsarbetet.
5. Om något måste backas:
   1. Kör `Restore-TeamsConfig ... -WhatIf` och granska ändringslistan (CSV-filen).
   2. Kör `Restore-TeamsConfig` utan `-WhatIf`, gärna begränsat med `-Include` eller `-User` till det som behöver återställas.
   3. Granska CSV-filen från den skarpa körningen efter rader med status `Misslyckad`.
   4. Jämför vid behov mot `BackupOnly/EffectivePolicies.csv` för att kontrollera att användarna har rätt policyer.
6. Ta gärna en ny backup efter ändringsarbetet och jämför den med den första med `Compare-TeamsConfigBackup`. Då syns allt som har ändrats, även sådant som inte var planerat.

## Backup-TeamsConfig

Läser ut konfigurationen och sparar den i en ny mapp med namnet `TeamsBackup_<tenant>_<ÅÅÅÅMMDD-TTMMSS>`.

```powershell
Backup-TeamsConfig [-Path <katalog>] [-VoiceUsersOnly] [-IncludeEffectivePolicies] [-IncludeUserCallingSettings] [-SkipUsers]
```

| Parameter | Beskrivning |
|---|---|
| `-Path` | Katalog där backupmappen skapas. Standard är aktuell katalog. |
| `-VoiceUsersOnly` | Läs bara ut användare med Enterprise Voice aktiverat. Snabbare i stora tenants. |
| `-IncludeEffectivePolicies` | Läs ut vilken policy som faktiskt gäller per användare, och om den kommer direkt eller via grupp. Tar ungefär 1 sekund per användare. Används för kontroll efter en återläsning. |
| `-IncludeUserCallingSettings` | Läs ut vidarekoppling, delegater och röstbrevlådeinställningar per röstanvändare. Tar ungefär 1–2 sekunder per användare. |
| `-SkipUsers` | Hoppa över användare och läs bara ut tenantkonfiguration. |

Om en enskild del inte går att läsa ut (till exempel för att en cmdlet saknas i modulversionen) fortsätter backupen med resten. Felet loggas i `backup.log` och sparas i `manifest.json`.

## Restore-TeamsConfig

Läser tillbaka konfiguration från en backupmapp. Skriptet jämför backupen med hur det ser ut i tenanten just nu och ändrar bara det som skiljer sig. Objekt som är oförändrade rörs inte.

```powershell
Restore-TeamsConfig -Path <backupmapp> [-Include <kategorier>] [-Exclude <kategorier>] [-User <upn>] [-PolicyType <typer>] [-RemoveExtra] [-WhatIf] [-Force] [-AllowDifferentTenant]
```

| Parameter | Beskrivning |
|---|---|
| `-Path` | Backupmappen, alltså den mapp som innehåller `manifest.json`. |
| `-Include` | En eller flera kategorier att läsa tillbaka. Standard är `All`. |
| `-Exclude` | Kategorier att hoppa över. |
| `-User` | Begränsa till en eller flera användare eller resurskonton (UPN eller ObjectId). Om `-Include` inte anges läses då bara `PhoneNumbers`, `Users` och `UserPolicies` tillbaka. |
| `-PolicyType` | Begränsa `Policies`, `GroupPolicyAssignments` och `UserPolicies` till vissa policytyper, till exempel `OnlineVoiceRoutingPolicy`, `TenantDialPlan` eller `TeamsCallingPolicy`. |
| `-RemoveExtra` | Ta även bort sådant som skapats efter backupen: SBC:er, voice routes, policyer, PSTN usages, grupptilldelningar och nummertilldelningar. Borttagningen körs sist, i omvänd ordning. |
| `-WhatIf` | Torrsimning. Visar alla kommandon som skulle köras, utan att ändra något. |
| `-Force` | Hoppa över bekräftelsefrågan innan start. |
| `-AllowDifferentTenant` | Tillåt återläsning till en annan tenant än den backupen togs från. Som standard avbryts körningen. |

### Kategorier

Kategorierna läses alltid tillbaka i den här ordningen, oavsett i vilken ordning de anges. Ordningen följer beroendena mellan objekten: en voice route kräver att SBC:n och PSTN usagen redan finns, en voice routing policy kräver att PSTN usagen finns, och så vidare.

| Ordning | Kategori | Innehåll |
|---|---|---|
| 1 | `PstnUsages` | PSTN usages |
| 2 | `TranslationRules` | Nummeröversättningsregler som används av SBC:erna |
| 3 | `Gateways` | SBC:er (Online PSTN Gateways) |
| 4 | `VoiceRoutes` | Voice routes |
| 5 | `VoiceRoutingPolicies` | Voice routing policies |
| 6 | `DialPlans` | Tenant dial plans inklusive normaliseringsregler |
| 7 | `Policies` | Övriga policydefinitioner (calling, caller ID, voicemail, nödsamtal, möten med flera) |
| 8 | `NetworkConfig` | Nätverksregioner, siter, subnät och betrodda IP-adresser |
| 9 | `NumberBlocking` | Blockerade och undantagna inkommande nummermönster |
| 10 | `GroupPolicyAssignments` | Policyer tilldelade via grupp |
| 11 | `PhoneNumbers` | Nummertilldelningar till användare och resurskonton |
| 12 | `Users` | Enterprise Voice och röstbrevlåda per användare |
| 13 | `UserPolicies` | Policyer tilldelade direkt till användare |

### Exempel

```powershell
$b = 'C:\Kund\Backup\TeamsBackup_Kund_20261018-090000'

# Bara routingen
Restore-TeamsConfig -Path $b -Include PstnUsages,Gateways,VoiceRoutes,VoiceRoutingPolicies

# Allt utom användarna
Restore-TeamsConfig -Path $b -Exclude PhoneNumbers,Users,UserPolicies

# Enskilda användare: nummer, Enterprise Voice och policyer
Restore-TeamsConfig -Path $b -User anna@kund.se,bertil@kund.se

# Bara voice routing policy och dial plan för alla användare
Restore-TeamsConfig -Path $b -Include UserPolicies -PolicyType OnlineVoiceRoutingPolicy,TenantDialPlan

# Exakt samma läge som vid backup, inklusive borttagning av nytillkomna objekt
Restore-TeamsConfig -Path $b -RemoveExtra -WhatIf
Restore-TeamsConfig -Path $b -RemoveExtra
```

### Telefonnummer

Om ett nummer har flyttats till en annan användare efter backupen tas det först bort från den nuvarande användaren och tilldelas sedan tillbaka. Om en användare har fått ett nytt nummer efter backupen tas det nya numret bort innan det ursprungliga tilldelas. De berörda användarna får ett kort avbrott medan detta pågår.

Nummer som var otilldelade vid backupen och som har tilldelats efteråt tas bara bort om `-RemoveExtra` anges.

## Torrsimning med -WhatIf

Med `-WhatIf` läser skriptet aktuell konfiguration, jämför mot backupen och visar varje kommando som skulle köras. Ingenting ändras i tenanten.

```
What if: Uppdatera OnlineVoiceRoute (OnlinePstnGatewayList, Priority)
    Set-CsOnlineVoiceRoute -Identity 'SE-National' -OnlinePstnGatewayList 'sbc1.kund.se','sbc2.kund.se' -Priority 1
What if: Tilldela OnlineVoiceRoutingPolicy till anna@kund.se (nu: VRP-Test)
    Grant-CsOnlineVoiceRoutingPolicy -Identity '3f2b...' -PolicyName 'VRP-Sverige'
```

I slutet visas antal planerade ändringar per kategori. Två filer sparas i backupmappen:

| Fil | Innehåll |
|---|---|
| `restore-<datum>-whatif.log` | Logg över körningen |
| `restore-<datum>-whatif.csv` | En rad per planerad ändring med kolumnerna Category, Change, Command och Status. Kan öppnas i Excel och användas som underlag för godkännande. |

Vid skarp körning skapas motsvarande `restore-<datum>.log` och `restore-<datum>.csv`. Där har varje ändring status `Utförd` eller `Misslyckad`, och felmeddelandet står i kolumnen Error.

Torrsimningen visar vad som kommer att göras, men inte om Teams kommer att acceptera det. Ett kommando som Teams avvisar, till exempel för att ett nummer inte längre finns i tenanten, märks först vid den skarpa körningen. Kör torrsimning och skarp körning nära i tid. Ändringar i tenanten mellan de två körningarna ger en annan plan.

## Get-TeamsConfigBackup

Visar information om en backup: tenant, datum, vem som tog den, eventuella fel samt antal objekt per fil.

```powershell
Get-TeamsConfigBackup -Path C:\Kund\Backup\TeamsBackup_Kund_20261018-090000 | Format-Table
```

Kolumnen `Restorable` visar om filen läses tillbaka automatiskt eller bara finns som referens.

## Compare-TeamsConfigBackup

Jämför två backuper och visar vad som skiljer. Kräver ingen anslutning till Teams, och ingenting ändras. Används till exempel för att se vad ett ändringsarbete faktiskt ändrade, eller för att ta reda på när och för vilka användare något försvann.

```powershell
Compare-TeamsConfigBackup [-ReferencePath] <backupmapp> [-DifferencePath] <backupmapp> [-Include <kategorier>] [-Exclude <kategorier>] [-User <upn>] [-PolicyType <typer>] [-Property <egenskaper>] [-CsvPath <fil>]
```

| Parameter | Beskrivning |
|---|---|
| `-ReferencePath` | Backupen att jämföra från, normalt den äldre. Dess värden visas i kolumnen `Before`. |
| `-DifferencePath` | Backupen att jämföra med, normalt den nyare. Dess värden visas i kolumnen `After`. |
| `-Include` | Kategorier att jämföra. Samma kategorier som för `Restore-TeamsConfig`, plus `BackupOnly` för filerna under `BackupOnly/`. Standard är `All`. |
| `-Exclude` | Kategorier att hoppa över. |
| `-User` | Begränsa till en eller flera användare eller resurskonton (UPN eller ObjectId). Om `-Include` inte anges jämförs då `PhoneNumbers`, `Users`, `UserPolicies` och användarfilerna under `BackupOnly` (`EffectivePolicies`, `UserCallingSettings`, `ResourceAccounts`). |
| `-PolicyType` | Begränsa `Policies`, `GroupPolicyAssignments` och `UserPolicies` till vissa policytyper. |
| `-Property` | Visa bara vissa egenskaper. Jokertecken tillåts, till exempel `LineUri`, `*VoiceRouting*` eller `CallingSettings.*`. |
| `-CsvPath` | Spara resultatet som semikolonseparerad CSV-fil för Excel. |

Resultatet är en rad per egenskap som skiljer sig:

| Kolumn | Innehåll |
|---|---|
| `Category` | Kategorin, till exempel `Users`, `PhoneNumbers` eller `VoiceRoutes` |
| `Type` | Objekttypen, till exempel `User`, `PhoneNumber`, `OnlineVoiceRoute` eller policytypen |
| `Item` | Objektet: UPN för användare, numret för telefonnummer, annars objektets namn |
| `Property` | Egenskapen som skiljer sig. Nästlade värden visas med punkt, till exempel `Policies.TeamsCallingPolicy`. |
| `Change` | `Tillagd`, `Borttagen` eller `Ändrad` |
| `Before`, `After` | Värdet i respektive backup |

Ett objekt som bara finns i den ena backupen, till exempel en användare som har tagits bort, ger en rad utan `Property`. När `-Property` anges visas i stället en rad per vald egenskap som har ett värde. Då kommer även borttagna användare med när man letar efter försvunna nummer.

### Exempel

```powershell
$fore  = 'C:\Kund\Backup\TeamsBackup_Kund_20261001-090000'
$efter = 'C:\Kund\Backup\TeamsBackup_Kund_20261018-090000'

# Alla skillnader, sparade för Excel
Compare-TeamsConfigBackup $fore $efter -CsvPath C:\Kund\diff.csv | Format-Table

# Användare som har tappat sitt telefonnummer
Compare-TeamsConfigBackup $fore $efter -Include Users -Property LineUri |
    Where-Object Change -eq 'Borttagen' | Format-Table Item, Before

# Alla nummer som har bytt eller tappat ägare
Compare-TeamsConfigBackup $fore $efter -Include PhoneNumbers -Property AssignedTo | Format-Table Item, Change, Before, After

# Allt som har ändrats för vissa användare
Compare-TeamsConfigBackup $fore $efter -User anna@kund.se,bertil@kund.se | Format-Table

# Ändrade voice routing policies, både definitioner och tilldelningar
Compare-TeamsConfigBackup $fore $efter -Include VoiceRoutingPolicies,UserPolicies -PolicyType OnlineVoiceRoutingPolicy | Format-Table

# Bläddra och filtrera i ett fönster (Windows)
Compare-TeamsConfigBackup $fore $efter | Out-GridView
```

### Att tänka på

- Telefonnummer finns på två ställen: `LineUri` på användaren (kategori `Users`) och nummertilldelningen (kategori `PhoneNumbers`, egenskapen `AssignedTo`). Kategorin `PhoneNumbers` täcker även resurskonton.
- En fil som bara finns i den ena backupen hoppas över med en varning. Annars skulle alla objekt i filen se tillagda eller borttagna ut. Det kan hända om backuperna togs med olika parametrar eller olika versioner av `MicrosoftTeams`-modulen.
- Om bara den ena backupen togs med `-VoiceUsersOnly` visas användare utan Enterprise Voice som tillagda eller borttagna. Funktionen varnar för detta.
- Tidsstämplar som ändras av sig själva (`WhenChanged`, `WhenCreated`, `LastSyncTimeStamp`) jämförs inte.
- Sammansatta värden, till exempel listor med objekt, jämförs som helhet och visas som JSON i `Before` och `After`.

## Vad som ingår i backupen

### Läses tillbaka automatiskt

| Fil | Innehåll |
|---|---|
| `DirectRouting/PstnUsages.json` | PSTN usages |
| `DirectRouting/TranslationRules.json` | Nummeröversättningsregler |
| `DirectRouting/Gateways.json` | SBC:er med alla inställningar (port, media bypass, failover, översättningsregler med mera) |
| `DirectRouting/VoiceRoutes.json` | Voice routes |
| `DirectRouting/VoiceRoutingPolicies.json` | Voice routing policies |
| `DirectRouting/DialPlans.json` | Dial plans med normaliseringsregler |
| `Network/*.json` | Regioner, siter, subnät och betrodda IP-adresser. Används av media bypass, Location-Based Routing och dynamiska nödsamtal. |
| `NumberBlocking/*.json` | Blockerade och undantagna inkommande nummermönster |
| `Policies/<Policytyp>.json` | En fil per policytyp. Vilka typer som tas med avgörs av vilka policyer som kan tilldelas användare i tenanten, plus en fast lista med röstrelaterade policyer. |
| `Assignments/PhoneNumberAssignments.json` | Alla telefonnummer i tenanten och vem de är tilldelade, inklusive resurskonton |
| `Assignments/GroupPolicyAssignments.json` | Policyer tilldelade via grupp, med prioritet |
| `Users/Users.json` | Användare med nummer, Enterprise Voice, röstbrevlåda och direkt tilldelade policyer |
| `Users/Users.csv` | Samma uppgifter som tabell, semikolonseparerad för Excel |

### Sparas endast som referens

Dessa filer ligger under `BackupOnly/` och läses inte tillbaka automatiskt. De används för manuell återställning eller kontroll.

| Fil | Innehåll | Varför ingen automatisk återläsning |
|---|---|---|
| `AutoAttendants.json`, `CallQueues.json` | Auto attendants och call queues | Får nya ID:n och tappar ljudfiler om de skapas på nytt |
| `ResourceAccounts.json` | Resurskonton | Kontona skapas i Entra ID. Numren på dem läses tillbaka via `PhoneNumbers`. |
| `OnlineSchedules.json` | Öppettider och helgdagar för auto attendants | Hör ihop med auto attendants |
| `Emergency/*.json` | Nödadresser och platser (LIS) samt subnät, switchar, portar och accesspunkter | Adress-ID:n ändras om de skapas på nytt |
| `SurvivableBranchAppliances.json` | Survivable Branch Appliances | Sällan förekommande, kräver manuell hantering |
| `Tenant/*.json` | Tenantinställningar: federation, uppgraderingsläge, klientinställningar, mötesinställningar, blockerade nummer, konferensinställningar | Tenantövergripande, bör ändras medvetet |
| `EffectivePolicies.json`, `.csv` | Vilken policy som faktiskt gällde per användare och varifrån den kom (direkt eller via grupp). Kräver `-IncludeEffectivePolicies`. | Uppgiften räknas fram av Teams och kan inte sättas direkt. Används för kontroll efter återläsning. |
| `UserCallingSettings.json` | Vidarekoppling, delegater och röstbrevlåda per användare. Kräver `-IncludeUserCallingSettings`. | Sparas som underlag för manuell återställning |

### Övriga filer

| Fil | Innehåll |
|---|---|
| `manifest.json` | Tenant-ID, tenantnamn, datum, vem som tog backupen, modulversion, antal objekt per fil och eventuella fel |
| `backup.log` | Logg från backupen |
| `restore-*.log`, `restore-*.csv` | Skapas vid varje återläsning eller torrsimning |

## Vad som inte ingår

- **Grupper och gruppmedlemskap.** De ligger i Entra ID. Grupptilldelningen av policyn sparas, men inte vilka användare som är med i gruppen.
- **Licenser.** Hanteras i Entra ID eller Microsoft 365 admin center.
- **Team, kanaler och innehåll.** Modulen hanterar bara telefoni- och policykonfiguration.
- **Konfiguration på själva SBC:erna.** Endast Teams-sidan av Direct Routing ingår. SBC:ns egen konfiguration måste säkerhetskopieras separat.
- **Certifikat och DNS.**

## Begränsningar och saker att tänka på

- **Kör alltid `-WhatIf` före en skarp återläsning.**
- **Användarnas policyfält visar bara direkt tilldelade policyer.** Policyer som kommer via grupp läses tillbaka via kategorin `GroupPolicyAssignments`.
- **Vissa policyinställningar har sammansatta värden** och läses inte tillbaka automatiskt, till exempel app-listor i `TeamsAppSetupPolicy`. Skriptet varnar om sådana värden skiljer sig mot backupen. Värdet finns i backupfilen och får sättas manuellt.
- **Inbyggda policyer** (till exempel `Global`) uppdateras men skapas eller tas aldrig bort.
- **Policyändringar kan ta upp till några timmar** innan de slår igenom i användarnas klienter.
- **Backupfilerna innehåller kunddata**: namn, UPN:er och telefonnummer. Förvara dem därefter. `.gitignore` i det här repot exkluderar `TeamsBackup_*` så att de inte checkas in av misstag.
- **Prova mot en testtenant** innan modulen används skarpt i en ny miljö eller efter en uppgradering av `MicrosoftTeams`-modulen.

## Felsökning

| Problem | Åtgärd |
|---|---|
| `Ingen anslutning till Teams` | Kör `Connect-MicrosoftTeams`. |
| `Backupen kommer från tenant ...` | Du är ansluten till fel tenant. Anslut till rätt tenant, eller använd `-AllowDifferentTenant` om det är avsiktligt. |
| `Kunde inte läsa ut ...` i backup.log | Cmdleten saknas eller misslyckades. Övriga delar av backupen påverkas inte. Uppdatera `MicrosoftTeams`-modulen om det gäller en cmdlet som saknas. |
| `... har ett sammansatt värde som inte återläses automatiskt` | Jämför värdet i backupfilen med nuvarande värde och sätt det manuellt vid behov. |
| Ändringar med status `Misslyckad` i restore-CSV:n | Felmeddelandet står i kolumnen Error. Kommandot i kolumnen Command kan köras manuellt efter åtgärd. |
