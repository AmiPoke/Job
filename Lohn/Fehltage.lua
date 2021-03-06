--------------------------------------------------------------------------------
--------------------------------------------------------------------------------
-- Betreff:    Gemeinsame Funktionalit�t der verschiedenen Lohnschnittstellen --
--                                                                            --
-- Author:     Jens Henske, ZMI GmbH                                          --
-- Erstellt:   17.08.2018                                                     --
-- Geaendert:                                                                 --
--                                                                            --
--------------------------------------------------------------------------------
--------------------------------------------------------------------------------


--############################################################################--
--###                                                                      ###--
--###                      � N D E R U N G E N                             ###--
--###                                                                      ###--
--###        Anpassung 28.08.2014: Umstellung auf Version 6                ###--
--###        Anpassung 21.10.2014: Ausfallschluessel(optional)             ###--
--###                                                                      ###--
--###                                                                      ###--
--############################################################################--

--############################################################################--
--###                                                                      ###--
--###                 H I L F S F U N K T I O N E N                        ###--
--###                                                                      ###--
--############################################################################--

-- Funktion zum Exportieren der Konten
function KontenExportieren(ID_MA_Stammdaten, ID_Lohnexport, f, Monat, Jahr, FirmenNr)
  -- zu exportierende Konten holen
  local k = SQLQuery:new('ZMI_Time','SELECT l.KontoNr, l.Lohnart, l.Wert, IFNULL(l.Dezimal, False) AS Dezimal, k.Zuschlag, l.Wert, k.Typ, ' ..
																		'IFNULL(k.Waehrung, False) AS Waehrung, IIF(UPPER(k.Info) = ' .. QuotedStr('BRUTTO') .. ', True, False) AS Brutto ' ..
																		'FROM Lohnexport_Konten l ' ..
																		'LEFT OUTER JOIN Konten k ON k.KontoNr = l.KontoNr ' ..
																		'WHERE l.ID_Lohnexport = ' .. ID_Lohnexport .. ' AND Wert<>1 AND l.Geloescht is null ' ..
																		'ORDER BY l.Dezimal')
	-- Alle Konten bis zum Ende duerchgehen
  while not k:Eof() do
		-- Variable zum  Pr�fen auf existierende Werte auf False setzen
    local bWerteVorhanden = false;
    -- Wert auf 0 setzen
		local Wert = 0
		--Personalnummer auf leer setzen
    local PersonalNr = ''
    -- Konten auf Dezimal False stellen
		local Dezimal = false
    -- KostenstellenNr auf leer setzen
		local KostenstellenNr = ''
		-- Kontonr ermitteln
    local KontoNr = k:FieldValue(0)
		-- Lohnart ermitteln
    local Lohnart = k:FieldValue(1)
		-- zuschlag ermitteln
    local Zuschlag = StrToFloat(k:FieldValue(4))
		-- W�erhung auf False setzen
    local Waehrung = false
		-- Brutto auf False setzen
    local Brutto = false
		-- WertTyp ermitteln
    local WertTyp = StrToInt(k:FieldValue(5))
		-- Pr�fen ob Zuschlag leer ist
    if Zuschlag == '' then
      -- Zuschlag auf NIL setzen
			Zuschlag = nil
    end
    -- Abfrage Konto auf Tabelle Ergebnis
    local sql = 'SELECT e.m, mas.Personalnummer FROM Ergebnis e LEFT OUTER JOIN MA_Stammdaten mas ON mas.ID = e.ID_MA_Stammdaten ' ..
							  'WHERE e.ID_MA_Stammdaten = ' .. ID_MA_Stammdaten ..
								' AND KontoNr = ' .. KontoNr ..
								' AND Monat = ' .. Monat .. ' AND Jahr = ' .. Jahr
		-- SQL ausf�hren
    local v = SQLQuery:new('ZMI_Time', sql)
		-- Pr�fen ob Datensatz vorhanden
    if (v:RecordCount() > 0) then
      -- Wert aus Datenbank ermitteln
			Wert = v:FieldValue(0)
      -- PersonalNr aus Datenbank ermitteln
			PersonalNr = v:FieldValue(1)
			-- Dezimal aus Datenbank ermitteln
      Dezimal = (k:FieldValue(3) == 'Wahr')
			-- Waehrung aus Datenbank ermitteln
      Waehrung = (k:FieldValue(7) == 'Wahr')
			-- Brutto aus Datenbank ermitteln
      Brutto = (k:FieldValue(8) == 'Wahr')
			-- KostenstellenNr leer zuweisen
      KostenstellenNr = ''
			-- Dat leer zuweisen
			Dat = ''
			-- Kostentraeger leer zuweisen
			Kostentraeger = ''
			-- Arbeitsart leer zuweisen
			Arbeitsart = ''
			-- LaText aus Datenbank ermitteln
			LaText = ''
			-- LGAText leer zuweisen
			LGAText = ''
			-- Vormonat ermitteln
			VMonat = Monat - 1
			-- VorJahr ermitteln
			VJahr = Jahr
			-- Pr�fen ob VMonat 0 ist
			if VMonat == 0 then
				-- VMonat auf 12 setzen
				VMonat = 12
				-- VJahr um ein Jahr zur�ck
				VJahr = Jahr - 1
			end

			-- Abfrage Gleitzeitkonto Vormonat
			qVM_GLZ = SQLQuery:new('ZMI_Time',	'SELECT ZMIF.MinToHour(mw.Jahresgleitzeit) FROM Monatswerte mw LEFT OUTER JOIN MA_Stammdaten mas ON mas.ID = mw.ID_MA_Stammdaten '..
																					'WHERE mw.ID_MA_Stammdaten = '..ID_MA_Stammdaten..' AND Monat = '..VMonat..' AND Jahr = '..VJahr)
			-- Pr�fen ob Datensatz vorhanden
			if qVM_GLZ:RecordCount() > 0 then
				LaText = LaText..'Gleitzeit Vormonat:'..qVM_GLZ:FieldValue(0)..', '
			end

			-- Abfrage aktuelles Gleitzeitkonto
			qAkt_GLZ = SQLQuery:new('ZMI_Time',	'SELECT ZMIF.MinToHour(mw.Jahresgleitzeit) FROM Monatswerte mw LEFT OUTER JOIN MA_Stammdaten mas ON mas.ID = mw.ID_MA_Stammdaten '..
																					'WHERE mw.ID_MA_Stammdaten = '..ID_MA_Stammdaten..' AND Monat = '..Monat..' AND Jahr = '..Jahr)
			-- Pr�fen ob Datensatz vorhanden
			if qAkt_GLZ:RecordCount() > 0 then
				LaText = LaText..'aktuelle Gleitzeit:'..qAkt_GLZ:FieldValue(0)
			end
			-- Variable zum  Pr�fen auf existierende Werte auf True setzen
      bWerteVorhanden = true;
    end
    -- Wete schreiben => Wenn Kontowert vorhanden
    if (bWerteVorhanden) then
      -- Aufruf der Funktion in der Lohnschnittstelle. Dort wird der String
			PersonaldatenKonten(FirmenNr, Monat, Jahr, PersonalNr,Dat,KostenstellenNr,Kostentraeger, Arbeitsart, Wert, Waehrung, Zuschlag, LaText, LGAText, Lohnart)
    end
		-- Daten sammeln
    collectgarbage();
    -- N�chster Datensatz
		k:Next()
  end
end

-- Funktion zum exportieren der Monatswerte
function MonatswerteExportieren(ID_MA_Stammdaten, ID_Lohnexport, f, Monat, Jahr, FirmenNr)
  -- Abfrage der zu exportierenenden Monatswerte
	local sql = 'SELECT Monatswert, Lohnart FROM Lohnexport_Monatswerte WHERE ID_Lohnexport = ' .. ID_Lohnexport .. ' AND Geloescht is null'
	-- Abfrage ausf�hren
  local m = SQLQuery:new('ZMI_Time', sql)
	-- Datens�tze durchgehen
  while not m:Eof() do
		-- Monatswert ermiteln
		local Monatswert = m:FieldValue(0)
		-- Lohnart ermiteln
    local Lohnart = m:FieldValue(1)
		-- Feld ermiteln
    local Feld = FelderMonatswerte[StrToInt(Monatswert)].Feld
		-- Dezimal ermiteln
    local Dezimal = FelderMonatswerte[StrToInt(Monatswert)].Dezimal
		-- Waehrung auf False setzen
    local Waehrung = false
		-- Brutto auf False setzen
    local Brutto = false
    -- Abfrage Personal
		local sql = 'SELECT Personalnummer p,ZMIF.MinToHour(m.Jahresgleitzeit), '.. Feld .. ' m ' .. 'FROM MA_Stammdaten mas LEFT OUTER JOIN Monatswerte m ON mas.ID = m.ID_MA_Stammdaten ' ..
								'WHERE m.Monat = ' .. Monat .. ' AND m.Jahr = ' .. Jahr ..' AND m.ID_MA_Stammdaten = ' .. ID_MA_Stammdaten
		-- Abfrage ausf�hren
    local w = SQLQuery:new('ZMI_Time', sql)
		-- Datens�tze durchgehen
    while not w:Eof() do
      -- PersonalNr ermitteln
			local PersonalNr = w:FieldValue(0)
			-- Vormonat ermitteln
			VMonat = Monat - 1
			-- VorJahr ermitteln
			VJahr = Jahr
			-- Pr�fen ob VMonat 0 ist
			if VMonat == 0 then
				-- VMonat auf 12 setzen
				VMonat = 12
				-- VJahr um ein Jahr zur�ck
				VJahr = Jahr - 1
			end
			LaText = ''
			-- Abfrage Gleitzeitkonto Vormonat
			qVM_GLZ = SQLQuery:new('ZMI_Time',	'SELECT ZMIF.MinToHour(mw.Jahresgleitzeit) FROM Monatswerte mw LEFT OUTER JOIN MA_Stammdaten mas ON mas.ID = mw.ID_MA_Stammdaten '..
																					'WHERE mw.ID_MA_Stammdaten = '..ID_MA_Stammdaten..' AND Monat = '..VMonat..' AND Jahr = '..VJahr)
			-- Pr�fen ob Datensatz vorhanden
			if qVM_GLZ:RecordCount() > 0 then
				-- LaText zuweisen
				LaText = LaText..'Gleitzeit Vormonat:'..qVM_GLZ:FieldValue(0)..', '
			end
			-- LaText zuweisen
			local	LaText = LaText..'aktuelle Gleitzeit:'..w:FieldValue(1)
      -- Wert ermitteln
			local Wert = w:FieldValue(2)
			-- KostenstellenNr Leer zuweisen
      local KostenstellenNr = ''
      -- Pr�fen ob Wert leer ist
      if Wert == '' then
        -- Wert auf 0 setzen
				Wert = '0'
      end
			-- Dat auf Leer setzen
			Dat = ''
			-- Kostentraeger auf Leer setzen
			Kostentraeger = ''
			-- Arbeitsart auf Leer setzen
			Arbeitsart = ''
			-- Waehrung auf Leer setzen
			Waehrung = ''
			-- Zuschlag auf Leer setzen
			Zuschlag = ''
			-- LGAText auf Leer setzen
			LGAText = ''
      -- Aufruf der Funktion im Lohnprogramm. Dort wird der String formatiert und in die Exportdatei geschrieben.
      PersonaldatenMonatswerte(FirmenNr, Monat, Jahr, PersonalNr,Dat,KostenstellenNr,Kostentraeger, Arbeitsart, Wert, Waehrung, Zuschlag, LaText, LGAText, Lohnart)
			-- N�chster Datensatz
      w:Next()
    end
    -- Query auf NIL setzen
		w = nil
		-- Daten Sammeln
    collectgarbage();
		-- N�chster Datensatz
    m:Next()
  end
end

--############################################################################--
--###                                                                      ###--
--###                 H A U P T F U N K T I O N E N                        ###--
--###                                                                      ###--
--############################################################################--

-- eigentliche Exportfunktion, regelt das File-Handle und aus ihr heraus werden die Kundenspezifischen Funktionen aufgerufen
function ExportData(MONAT, JAHR, ID_Lohnexport, Dateiname)
  -- zu exportierende Konten aus Datenbank auslesen
  local k = SQLQuery:new('ZMI_Time','SELECT KontoNr, Lohnart, Wert FROM Lohnexport_Konten WHERE ID_Lohnexport = ' .. ID_Lohnexport .. ' AND Geloescht IS NULL')
	-- Logeintrag erzeugen
  DebugMessage('Es werden ' .. k:RecordCount() .. ' Konten exportiert.')
  -- zu exportierende Monatswerte aus Datenbank auslesen
  local q = SQLQuery:new('ZMI_Time','SELECT Monatswert, Lohnart FROM Lohnexport_Monatswerte WHERE ID_Lohnexport = ' .. ID_Lohnexport .. ' AND Geloescht IS NULL')
  -- Logeintrag erzeugen
  DebugMessage('Es werden ' .. q:RecordCount() .. ' Monatswerte exportiert.')
  -- Mitarbeiter aus Datenbank ermitteln
  local p = SQLQuery:new('ZMI_Time','SELECT mas.ID, f.FirmenNr FROM MA_Stammdaten mas Left Outer Join Firma f ON f.ID = mas.ID_Firma WHERE mas.ID_Lohnexport = ' .. ID_Lohnexport .. ' AND IFNULL(mas.Passiv, false) = false AND '..
																		' (mas.Austritt>='..QuotedStr(LastDay(MONAT, JAHR)..'.'..MONAT..'.'..JAHR)..' OR mas.Austritt IS NULL) AND mas.Geloescht IS NULL')
	-- Logeintrag erzeugen
  DebugMessage('Es werden ' .. p:RecordCount() .. ' Personaldaten exportiert.')
	-- Pfad ermitteln
  local ExportFile = GetPath(ID_Lohnexport, Dateiname)
  -- Pr�fen ob File ge�ffnet wurde
  if FileExists(ExportFile) then
	  -- Abfrage zum �berschreiben des Files erzeugen
    local res = MessageBox('Datei ' .. ExportFile .. ' existiert bereits. Wollen Sie die bestehende Datei l�schen?', 'Datei existiert bereits',{MB_YESNOCANCEL, MB_ICONQUESTION})
		-- Pr�fen ob JA geklickt wurde
    if res == IDYES then
			-- File l�schen
      DeleteFile(ExportFile)
			-- File erzeugen
      f = CreateFile(ExportFile)
			-- Pr�fen ob File erzeugt wurde
      if not f then
        ShowMessage('Exportdatei ' .. ExportFile .. ' konnte nicht ' ..
          'erstellt werden. Bitte pr�fen Sie die Pfadangaben')
      end
		-- Pr�fen ob Nein geklickt wurde
    elseif res == IDNO then
      -- File �ffnen
			f = OpenFile(ExportFile)
      -- Ans Ende der File springen
			SeekFile(f, 0, 'end')
    else
			-- Error erzeugen, dass Export abgebrochen
      error('Export abgebrochen')
    end
  else
		-- File erzeugen
    f = CreateFile(ExportFile)
		-- Pr�fen ob File ge�ffnet
    if not f then
			-- Message erzeugen
      ShowMessage('Exportdatei konnte nicht erstellt werden. Bitte pr�fen Sie die Pfadangaben')
    end
  end
  -- Pr�fen ob File ge�ffnet
  if f then
    -- Ab hier beginnt der Mitarbeiterbezogene Export der Datens�tze
    while not p:Eof() do
			-- ID des Mitarbeiters ermitteln
      ID_MA_Stammdaten = p:FieldValue(0)
			-- MandantenNr ermitteln
			MandantenNr = p:FieldValue(1)
      -- Konten Monat exportieren
      KontenExportieren(ID_MA_Stammdaten, ID_Lohnexport, f, MONAT, JAHR, MandantenNr)
      -- Monatswerte exportieren
      MonatswerteExportieren(ID_MA_Stammdaten, ID_Lohnexport, f, MONAT, JAHR, MandantenNr)
      -- n�chster Datensatz
			p:Next()
    end
		-- File schliessen
    CloseFile(f)
  end
end
