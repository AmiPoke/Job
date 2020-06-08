--------------------------------------------------------------------------------
--------------------------------------------------------------------------------
--                                                                            --
-- Betreff:    Monatsliste Eintritt                                           --
--                                                                            --
-- Author:     Jens Henske, ZMI GmbH                                          --
-- Erstellt:   14.08.2018                                                     --
-- Geaendert:  																															  --
--                                                                            --
--------------------------------------------------------------------------------
--------------------------------------------------------------------------------

--############################################################################--
--###                                                                      ###--
--###                      Ä N D E R U N G E N                             ###--
--###                                                                      ###--
--############################################################################--

--############################################################################--
--###                                                                      ###--
--###                        I N C L U D E S                               ###--
--###                                                                      ###--
--############################################################################--

-- include
Include('config.lua')

--############################################################################--
--###                                                                      ###--
--###                       V A R I A B L E N                              ###--
--###                                                                      ###--
--############################################################################--

--############################################################################--
--###                                                                      ###--
--###                 H I L F S F U N K T I O N E N                        ###--
--###                                                                      ###--
--############################################################################--

-- Funktion zum ermitteln des Pfades und Dateiname
function GetPath(ID_Lohnexport, Dateiname)
	-- Abfrage Pfad und Dateiname 
  q = SQLQuery:new('ZMI_Time', 'SELECT Exportpfad, Dateiname FROM Lohnexport WHERE ID = ' .. ID_Lohnexport)
  -- Prüfen ob Datensatz vorhanden
	if q then
    -- Pfad auslesen
		Path = IncludeTrailingBackslash(q:FieldValue(0))
    -- Prüfen auf Dateiname
    if Dateiname then
			-- Dateiname festlegen
      Name = Dateiname
    else
			-- Dateiname festlegen
      Name = q:FieldValue(1)
    end
		-- Pfad und Dateiname als Wert zurückliefern
    return Path .. Name
  else
		-- leeren Wert zurückliefern
		return ''
  end
end

--############################################################################--
--###                                                                      ###--
--###                 H A U P T F U N K T I O N E N                        ###--
--###                                                                      ###--
--############################################################################--

function ExportData()
	-- Abfrage letzter Tag im Monat
	l = SQLQuery:new('ZMI_Time','SELECT ZMIF.GetLastDayOfMonth('..JAHR..','..MONAT..') FROM system.iota;')
	-- Letzten Tag aus Datenbank auslesen
	sLetzterTag = l:FieldValue(0)
	-- Datum für Prüfung zusammenstellen
	CheckDatumEintritt = sLetzterTag..'.'..MONAT..'.'..JAHR 
	CheckDatumAustritt = '01.'..MONAT..'.'..JAHR
	-- Abfrage aller Mitarbeiter die im Monat und Jahr eingetreten sind
	p = SQLQuery:new('ZMI_Time',"SELECT mas.Nachname, mas.Vorname, maa.Strasse, maa.Plz, maa.Ort "..
																		"From Ma_Stammdaten mas "..
																		"Left Outer Join MA_Adressen maa ON maa.ID_MA_Stammdaten = mas.id "..
																		"Where mas.Eintritt <= "..QuotedStr(CheckDatumEintritt)..
																		" AND IFNULL(mas.Austritt,Cast('31.12.2999' as sql_Date))  >= "..QuotedStr(CheckDatumAustritt))
	-- Exportfile Ermitteln
	ExportFile = GetPath(ID_Lohnexport, Dateiname)
	-- Prüfen ob Datei existriert
	if FileExists(ExportFile) then
		-- File löschen 
		DeleteFile(ExportFile)
		-- File erzeugen
		f = CreateFile(ExportFile)
	else
		-- File erzeugen
		f = CreateFile(ExportFile)
	end
	-- Prüfen ob File geöffnet werden kann
	if f then
		-- Header Zusammenstellen
		Header = 'Name, Vorname; Straße mit Hausnummer;PLZ;Ort'
		-- Header in Datei schreiben
		-- WriteLine(f, Header)
		-- Datensätze durchgehen
		while not p:Eof() do
			-- Werte in Datei schreiben
			WriteLine(f,';;'..p:FieldValue(0)..', '..p:FieldValue(1)..';'..p:FieldValue(2)..';'..p:FieldValue(3)..';'..p:FieldValue(4)..';')
			-- Nächster Datensatz
			p:Next()
		end
		CloseFile(f)
	end
end

--############################################################################--
--###                                                                      ###--
--###                  P R O G R A M M A N F A N G                         ###--
--###                                                                      ###--
--############################################################################--

-- Exportfunktion aufrufen
ExportData()
ShowMessage('Daten wurden erfolgreich exportiert!')

--############################################################################--
--###                                                                      ###--
--###                     P R O G R A M M E N D E                          ###--
--###                                                                      ###--
--############################################################################--
