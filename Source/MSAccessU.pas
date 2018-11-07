unit MSAccessU;

interface

type
  TByteArray = array[0..85] of byte;

  function XorPassword(Bytes: TByteArray): String;

implementation

function XorPassword(Bytes: TByteArray): String;
const
    XorBytes: array[0..17] of byte = ($86, $FB, $EC, $37, $5D, $44, $9C, $FA, $C6, $5E, $28, $E6, $13, $B6, $8A, $60, $54, $94);
var
   i: Integer;
   CurrChar: Char;
begin
    Result := '';
    for i := 0 to 17 do begin
        CurrChar := chr(ord(bytes[i + $42]) xor XorBytes[i]);
        if CurrChar = #0 then
           break;
        Result := Result + CurrChar;
    end;
end;

end.
