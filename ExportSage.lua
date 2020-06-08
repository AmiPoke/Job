--------------------------------------------------------------------------------
--------------------------------------------------------------------------------
--                                                                            --
-- Betreff:    Lohnexport zu Sage                                             --
--                                                                            --
-- Author:     Jens Henske, ZMI GmbH                                          --
-- Erstellt:   17.08.2018                                                     --
-- Geaendert:  																										            --
--                                                                            --
--------------------------------------------------------------------------------
--------------------------------------------------------------------------------

Include('config.lua');
Include('lohn\\LohnKombiV6.lua')
Include('lohn\\ZMI.lua')

--############################################################################--
--###                                                                      ###--
--###                      Ä N D E R U N G E N                             ###--
--###                                                                      ###--
--############################################################################--

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

--############################################################################--
--###                                                                      ###--
--###                 H A U P T F U N K T I O N E N                        ###--
--###                                                                      ###--
--############################################################################--


function WerteSchreiben(FirmenNr, Monat, Jahr, PersonalNr,Dat,KostenstellenNr,Kostentraeger, Arbeitsart, Wert, Waehrung, Zuschlag, LaText, LGAText,Lohnart)
	-- Prüefn ob Wert <> 0 
  if StrToFloat(Wert) ~= 0 then
		-- Prüfen ob Dezimal 
		if not Dezimal then
			-- Wert umwandeln
			Wert = IntTimeToStr(StrToFloat(Wert))
		else
			-- Wert umwandeln
			Wert = FormatStr(StrToFloat(Wert))
		end
		-- Waehrung auf Leer setzen
		Waehrung = ''
		-- Zuschlag auf Leer setzen
		Zuschlag = ''
		-- LAText auf leer setzen (gibt die Gleitzeit aus, bei Bedarf auskommentieren)
		LGAText = LaText
		LaText = ''
    -- Line zusammenstellen
	  Line = FirmenNr..';'..Monat..';'..Jahr..';'..PersonalNr..';'..Lohnart..';'..Dat..';'..KostenstellenNr..';'..Kostentraeger..';'..Arbeitsart..';'..Wert..';'..Waehrung..';'..Zuschlag..';'..LaText..';'..LGAText
			
    WriteLine(f, Line)
  end
end

-- Personaldaten Konten schreiben
function PersonaldatenKonten(FirmenNr, Monat, Jahr, PersonalNr,Dat,KostenstellenNr,Kostentraeger, Arbeitsart, Wert, Waehrung, Zuschlag, LaText, LGAText,Lohnart)
  WerteSchreiben(FirmenNr, Monat, Jahr, PersonalNr,Dat,KostenstellenNr,Kostentraeger, Arbeitsart, Wert, Waehrung, Zuschlag, LaText, LGAText,Lohnart)
end

-- Personaldaten Monatswerte schreiben
function PersonaldatenMonatswerte(FirmenNr, Monat, Jahr, PersonalNr,Dat,KostenstellenNr,Kostentraeger, Arbeitsart, Wert, Waehrung, Zuschlag, LaText, LGAText,Lohnart)
  WerteSchreiben(FirmenNr, Monat, Jahr, PersonalNr,Dat,KostenstellenNr,Kostentraeger, Arbeitsart, Wert, Waehrung, Zuschlag, LaText, LGAText,Lohnart)
end

--############################################################################--
--###                                                                      ###--
--###                  P R O G R A M M A N F A N G                         ###--
--###                                                                      ###--
--############################################################################--

-- Exportfunktion aufrufen
ExportData(MONAT, JAHR, ID_Lohnexport)

ShowMessage('Daten wurden erfolgreich exportiert!')

--############################################################################--
--###                                                                      ###--
--###                     P R O G R A M M E N D E                          ###--
--###                                                                      ###--
--############################################################################--
