-- Letzte Änderung:
-- 22.07.2016 Christian Erhard, ZMI GmbH

FelderMonatswerte = {
  {Feld = 'IststundenBrutto', Dezimal = false},
  {Feld = 'IststundenNetto', Dezimal = false},
  {Feld = 'Sollstunden', Dezimal = false},
  {Feld = 'Mo_gleit', Dezimal = false},
  {Feld = 'Ausbezahlt', Dezimal = false},
  {Feld = 'Jahresgleitzeit', Dezimal = false},
  {Feld = 'UrlaubstageBezahlt', Dezimal = true},
  {Feld = 'UrlaubsstundenBezahlt', Dezimal = false},
  {Feld = 'KrankentageBezahlt', Dezimal = true},
  {Feld = 'KrankstundenBezahlt', Dezimal = false},
  {Feld = 'KrankentageUnbezahlt', Dezimal = true},
  {Feld = 'KrankstundenUnbez', Dezimal = false},
  {Feld = 'Feier_tage', Dezimal = true},
  {Feld = 'Feiertagsstunden', Dezimal = false},
  {Feld = 'Berufsschule', Dezimal = false},
  {Feld = 'Freischicht', Dezimal = true},
  {Feld = 'Dienstreise', Dezimal = true},
  {Feld = 'Resturlaub', Dezimal = true},
  {Feld = 'Kappungskonto', Dezimal = false},
  {Feld = 'KappungsKMonats_GL', Dezimal = false},
  {Feld = 'KappungsKJahres_GL', Dezimal = false}
}

function GetKW(ADatum)
	query_Woche = SQLQuery:new('ZMI_Time','Select ISOWEEK('..QuotedStr(ADatum)..') from System.iota')
	
	local Kalenderwoche = 0
	
	if (query_Woche:RecordCount() > 0) then
		Kalenderwoche = AddZeros(query_Woche:FieldValue(0), 2)
	end
		
	return Kalenderwoche
end;

function GetPath(ID_Lohnexport, Dateiname)
  local q = SQLQuery:new('ZMI_Time', 'SELECT Exportpfad, Dateiname FROM Lohnexport WHERE ID = ' .. ID_Lohnexport)

  if q then
    local Path = IncludeTrailingBackslash(q:FieldValue(0))
    local Name

    if Dateiname then
        Name = Dateiname
    else
        Name = q:FieldValue(1)
    end

    return Path .. Name
  else
    return ''
  end
end

function GetPathOnly(ID_Lohnexport, Dateiname)
  local q = SQLQuery:new('ZMI_Time', 'SELECT Exportpfad, Dateiname FROM Lohnexport WHERE ID = ' .. ID_Lohnexport)

  if q then
    local Path = IncludeTrailingBackslash(q:FieldValue(0))
    local Name

    if Dateiname then
        Name = Dateiname
    else
        Name = q:FieldValue(1)
    end

    return Path
  else
    return ''
  end
end

function AddZeros(s, count)
  if s == nil then
    DebugMessage('AddZeros mit (nil, ' .. count .. ') aufgerufen')
  end

  if Length(s) > count then
    -- die letzten count Stellen kopieren
    s = Copy(s, Length(s) - count + 1, count);
  else
    -- '0'-en hinzufügen
    while Length(s) < count do
      s = '0' .. s
    end
  end

  return s
end

function AddZeros2(s, count)
  if s == nil then
    DebugMessage('AddZeros mit (nil, ' .. count .. ') aufgerufen')
  end

  if Length(s) > count then
    -- die letzten count Stellen kopieren
    s = Copy(s, Length(s) - count + 1, count);
  else
    -- '0'-en hinzufügen
    while Length(s) < count do
      s = s .. '0'
    end
  end

  return s
end

function AddEmptyString(s, count)
  if Length(s) > count then
    -- die letzten count Stellen kopieren
    s = Copy(s, Length(s) - count + 1, count);
  else
    -- '0'-en hinzufügen
    while Length(s) < count do
      s = s .. ' '
    end
  end

  return s
end

function AddEmptyString2(s, count)
  if Length(s) > count then
    -- die letzten count Stellen kopieren
    s = Copy(s, Length(s) - count + 1, count);
  else
    -- '0'-en hinzufügen
    while Length(s) < count do
      s = ' ' .. s
    end
  end

  return s
end

function IntTimeToStr(value)
  value = value / 60
  nK = Round(Frac(value) * 100)

  if (nK < 0) then
    nK = nK * -1
  end

  vorKomma = 0

  if (value < 0 and Trunc(value) == 0) then
    vorKomma = '-'..Trunc(value)
  else
    vorKomma = Trunc(value)
  end

  --Wert mit 2 Nachkommastellen zurckgeben
  return vorKomma .. "," .. AddZeros(Trunc(nK), 2)
end

function LastDay(monat, jahr)
  -- 1. des nächsten Monats ermitteln und 1 abziehen
  if monat == 12 then
    monat = 1
    jahr = jahr + 1
  else
    monat = monat + 1
  end

  local d = EncodeDate(jahr, monat, 1) - 1
  local jahr, monat, tag = DecodeDate(d)

  return tag
end

function WertVorhanden(sWert)
  return StrToFloat(sWert) ~= 0
end

function FormatStr(value)
  --Wert mit 2 Nachkommastellen zurückgeben
  return Trunc(value) .. "," .. AddZeros(Trunc(Frac(value) * 100), 2)
end

function FormatStrMW(value)
  --Wert mit 2 Nachkommastellen zurückgeben
  local iPos = Pos(',', value)
  local fValue = ''
  if iPos > 0 then
    return Copy(value, 1, iPos - 1) .. ',' .. AddZeros2(Copy(value, iPos + 1, 2), 2)
  else
    return value .. ',' .. AddZeros('', 2)
  end
  
 -- return Trunc(value) .. "," .. AddZeros(Trunc(Frac(value) * 100), 2)
end

function GetTableValue(Table, Index)
  erg = 'N/A'

  for key, value in Table do
    if(value[1] == Index) then
      erg = value[2]
   end
  end
  return erg
end

function GetShortDayOfWeek(iWochentag)
  local sWochentag = ''

  if iWochentag == 1 then
    sWochentag = "So"
  elseif iWochentag == 2 then
    sWochentag = "Mo"
  elseif iWochentag == 3 then
    sWochentag = "Di"
  elseif iWochentag == 4 then
    sWochentag = "Mi"
  elseif iWochentag == 5 then
    sWochentag = "Do"
  elseif iWochentag == 6 then
    sWochentag = "Fr"
  elseif iWochentag == 7 then
    sWochentag = "Sa"
  end;

  return sWochentag
end