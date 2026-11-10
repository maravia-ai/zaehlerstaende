// Konfiguration der Zählerstände-App.
// Nach der App-Registrierung in Microsoft Entra ID hier die beiden IDs eintragen.
window.APP_CONFIG = {
  // "Anwendungs-ID (Client)" aus der Entra-App-Registrierung
  clientId: "HIER_CLIENT_ID_EINTRAGEN",

  // "Verzeichnis-ID (Mandant)" aus der Entra-App-Registrierung (gleich wie bei der Medikamente-App)
  tenantId: "e4e57d17-93c0-4e0e-a004-ec2c9ea7a8e0",

  // Pfad der Excel-Datei in deinem OneDrive (relativ zum OneDrive-Stammordner)
  filePath: "1 - Dirk/06 Immobilien/Zählerstände.xlsx",

  // Ordner für die Zählerfotos (wird bei Bedarf angelegt)
  fotoOrdner: "1 - Dirk/06 Immobilien/Zählerstände Fotos",

  // Monatliche Kalender-Erinnerung: Tag im Monat und Uhrzeit
  erinnerungTag: 1,
  erinnerungUhrzeit: "09:00",

  // Zeitzone für Kalendertermine (Windows-Bezeichnung)
  zeitzone: "W. Europe Standard Time",

  // Gebäude und Zähler. Reihenfolge = Reihenfolge in der App.
  // Die Zählernummern stehen im Blatt "Zähler" der Excel-Datei.
  gebaeude: ["Burgacker 39", "Leipziger Str. 7"],
  medien: [
    { name: "Gas",    einheit: "m³",  nachkomma: 3 },
    { name: "Wasser", einheit: "m³",  nachkomma: 3 },
    { name: "Strom",  einheit: "kWh", nachkomma: 1 }
  ]
};
