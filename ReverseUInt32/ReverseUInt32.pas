// Function to reverse the bits of a UInt32
//
// Basically, the idea is to swap 2 bits, then 4 bits, 8 bits and 16 bits
// If an integer binary representation is (abcdefgh) then what the code is:
//
//    abcdefgh - ba dc fe hg - dcba hgfe - hgfedcba

function ReverseBits(const Value: UInt32): UInt32;
begin
  Result := Value;

  Result := ((Result shr  1) and $55555555) or ((Result shl  1) and $AAAAAAAA);
  Result := ((Result shr  2) and $33333333) or ((Result shl  2) and $CCCCCCCC);
  Result := ((Result shr  4) and $0F0F0F0F) or ((Result shl  4) and $F0F0F0F0);
  Result := ((Result shr  8) and $00FF00FF) or ((Result shl  8) and $FF00FF00);
  Result := ((Result shr 16) and $0000FFFF) or ((Result shl 16) and $FFFF0000);
end;
