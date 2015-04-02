function _JS2ANSI(str) { if(typeof cptable !== 'undefined') return cptable.utils.encode(1252, str); return str.split("").map(function(x) { return x.charCodeAt(0); }); }

/* [MS-OFFCRYPTO] 2.1.4 Version */
function parse_Version(blob, length) {
	var o = {};
	o.Major = blob.read_shift(2);
	o.Minor = blob.read_shift(2);
	return o;
}
/* [MS-OFFCRYPTO] 2.3.2 Encryption Header */
function parse_EncryptionHeader(blob, length) {
	var o = {};
	o.Flags = blob.read_shift(4);

	// Check if SizeExtra is 0x00000000
	var tmp = blob.read_shift(4);
	if(tmp !== 0) throw 'Unrecognized SizeExtra: ' + tmp;

	o.AlgID = blob.read_shift(4);
	switch(o.AlgID) {
		case 0: case 0x6801: case 0x660E: case 0x660F: case 0x6610: break;
		default: throw 'Unrecognized encryption algorithm: ' + o.AlgID;
	}
	parsenoop(blob, length-12);
	return o;
}

/* [MS-OFFCRYPTO] 2.3.3 Encryption Verifier */
function parse_EncryptionVerifier(blob, length) {
	return parsenoop(blob, length);
}
/* [MS-OFFCRYPTO] 2.3.5.1 RC4 CryptoAPI Encryption Header */
function parse_RC4CryptoHeader(blob, length) {
	var o = {};
	var vers = o.EncryptionVersionInfo = parse_Version(blob, 4); length -= 4;
	if(vers.Minor != 2) throw 'unrecognized minor version code: ' + vers.Minor;
	if(vers.Major > 4 || vers.Major < 2) throw 'unrecognized major version code: ' + vers.Major;
	o.Flags = blob.read_shift(4); length -= 4;
	var sz = blob.read_shift(4); length -= 4;
	o.EncryptionHeader = parse_EncryptionHeader(blob, sz); length -= sz;
	o.EncryptionVerifier = parse_EncryptionVerifier(blob, length);
	return o;
}
/* [MS-OFFCRYPTO] 2.3.6.1 RC4 Encryption Header */
function parse_RC4Header(blob, length) {
	var o = {};
	var vers = o.EncryptionVersionInfo = parse_Version(blob, 4); length -= 4;
	if(vers.Major != 1 || vers.Minor != 1) throw 'unrecognized version code ' + vers.Major + ' : ' + vers.Minor;
	o.Salt = blob.read_shift(16);
	o.EncryptedVerifier = blob.read_shift(16);
	o.EncryptedVerifierHash = blob.read_shift(16);
	return o;
}

/* [MS-OFFCRYPTO] 2.3.7.1 Binary Document Password Verifier Derivation */
function crypto_CreatePasswordVerifier_Method1(Password) {
	var Verifier = 0x0000, PasswordArray;
	var PasswordDecoded = _JS2ANSI(Password);
	var len = PasswordDecoded.length + 1, i, PasswordByte;
	var Intermediate1, Intermediate2, Intermediate3;
	PasswordArray = new_raw_buf(len);
	PasswordArray[0] = PasswordDecoded.length;
	for(i = 1; i != len; ++i) PasswordArray[i] = PasswordDecoded[i-1];
	for(i = len-1; i >= 0; --i) {
		PasswordByte = PasswordArray[i];
		Intermediate1 = ((Verifier & 0x4000) === 0x0000) ? 0 : 1;
		Intermediate2 = (Verifier << 1) & 0x7FFF;
		Intermediate3 = Intermediate1 | Intermediate2;
		Verifier = Intermediate3 ^ PasswordByte;
	}
	return Verifier ^ 0xCE4B;
}

/* [MS-OFFCRYPTO] 2.3.7.2 Binary Document XOR Array Initialization */
var crypto_CreateXorArray_Method1 = (function() {
	var PadArray = [0xBB, 0xFF, 0xFF, 0xBA, 0xFF, 0xFF, 0xB9, 0x80, 0x00, 0xBE, 0x0F, 0x00, 0xBF, 0x0F, 0x00];
	var InitialCode = [0xE1F0, 0x1D0F, 0xCC9C, 0x84C0, 0x110C, 0x0E10, 0xF1CE, 0x313E, 0x1872, 0xE139, 0xD40F, 0x84F9, 0x280C, 0xA96A, 0x4EC3];
	var XorMatrix = [0xAEFC, 0x4DD9, 0x9BB2, 0x2745, 0x4E8A, 0x9D14, 0x2A09, 0x7B61, 0xF6C2, 0xFDA5, 0xEB6B, 0xC6F7, 0x9DCF, 0x2BBF, 0x4563, 0x8AC6, 0x05AD, 0x0B5A, 0x16B4, 0x2D68, 0x5AD0, 0x0375, 0x06EA, 0x0DD4, 0x1BA8, 0x3750, 0x6EA0, 0xDD40, 0xD849, 0xA0B3, 0x5147, 0xA28E, 0x553D, 0xAA7A, 0x44D5, 0x6F45, 0xDE8A, 0xAD35, 0x4A4B, 0x9496, 0x390D, 0x721A, 0xEB23, 0xC667, 0x9CEF, 0x29FF, 0x53FE, 0xA7FC, 0x5FD9, 0x47D3, 0x8FA6, 0x0F6D, 0x1EDA, 0x3DB4, 0x7B68, 0xF6D0, 0xB861, 0x60E3, 0xC1C6, 0x93AD, 0x377B, 0x6EF6, 0xDDEC, 0x45A0, 0x8B40, 0x06A1, 0x0D42, 0x1A84, 0x3508, 0x6A10, 0xAA51, 0x4483, 0x8906, 0x022D, 0x045A, 0x08B4, 0x1168, 0x76B4, 0xED68, 0xCAF1, 0x85C3, 0x1BA7, 0x374E, 0x6E9C, 0x3730, 0x6E60, 0xDCC0, 0xA9A1, 0x4363, 0x86C6, 0x1DAD, 0x3331, 0x6662, 0xCCC4, 0x89A9, 0x0373, 0x06E6, 0x0DCC, 0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48C4];
	var Ror = function(Byte) { return ((Byte/2) | (Byte*128)) & 0xFF; };
	var XorRor = function(byte1, byte2) { return Ror(byte1 ^ byte2); };
	var CreateXorKey_Method1 = function(Password) {
		var XorKey = InitialCode[Password.length - 1];
		var CurrentElement = 0x68;
		for(var i = Password.length-1; i >= 0; --i) {
			var Char = Password[i];
			for(var j = 0; j != 7; ++j) {
				if(Char & 0x40) XorKey ^= XorMatrix[CurrentElement];
				Char *= 2; --CurrentElement;
			}
		}
		return XorKey;
	};
	return function(password) {
		var Password = _JS2ANSI(password);
		var XorKey = CreateXorKey_Method1(Password);
		var Index = Password.length;
		var ObfuscationArray = new_raw_buf(16);
		for(var i = 0; i != 16; ++i) ObfuscationArray[i] = 0x00;
		var Temp, PasswordLastChar, PadIndex;
		if((Index & 1) === 1) {
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(PadArray[0], Temp);
			--Index;
			Temp = XorKey & 0xFF;
			PasswordLastChar = Password[Password.length - 1];
			ObfuscationArray[Index] = XorRor(PasswordLastChar, Temp);
		}
		while(Index > 0) {
			--Index;
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
			--Index;
			Temp = XorKey & 0xFF;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
		}
		Index = 15;
		PadIndex = 15 - Password.length;
		while(PadIndex > 0) {
			Temp = XorKey >> 8;
			ObfuscationArray[Index] = XorRor(PadArray[PadIndex], Temp);
			--Index;
			--PadIndex;
			Temp = XorKey & 0xFF;
			ObfuscationArray[Index] = XorRor(Password[Index], Temp);
			--Index;
			--PadIndex;
		}
		return ObfuscationArray;
	};
})();

/* [MS-OFFCRYPTO] 2.3.7.3 Binary Document XOR Data Transformation Method 1 */
var crypto_DecryptData_Method1 = function(password, Data, XorArrayIndex, XorArray, O) {
	/* If XorArray is set, use it; if O is not set, make changes in-place */
	if(!O) O = Data;
	if(!XorArray) XorArray = crypto_CreateXorArray_Method1(password);
	var Index, Value;
	for(Index = 0; Index != Data.length; ++Index) {
		Value = Data[Index];
		Value ^= XorArray[XorArrayIndex];
		Value = ((Value>>5) | (Value<<3)) & 0xFF;
		O[Index] = Value;
		++XorArrayIndex;
	}
	return [O, XorArrayIndex, XorArray];
};

var crypto_MakeXorDecryptor = function(password) {
	var XorArrayIndex = 0, XorArray = crypto_CreateXorArray_Method1(password);
	return function(Data) {
		var O = crypto_DecryptData_Method1(null, Data, XorArrayIndex, XorArray);
		XorArrayIndex = O[1];
		return O[0];
	};
};

/* 2.5.343 */
function parse_XORObfuscation(blob, length, opts, out) {
	var o = { key: parseuint16(blob), verificationBytes: parseuint16(blob) };
	if(opts.password) o.verifier = crypto_CreatePasswordVerifier_Method1(opts.password);
	out.valid = o.verificationBytes === o.verifier;
	if(out.valid) out.insitu_decrypt = crypto_MakeXorDecryptor(opts.password);
	return o;
}

/* 2.4.117 */
function parse_FilePassHeader(blob, length, oo) {
	var o = oo || {}; o.Info = blob.read_shift(2); blob.l -= 2;
	if(o.Info === 1) o.Data = parse_RC4Header(blob, length);
	else o.Data = parse_RC4CryptoHeader(blob, length);
	return o;
}
function parse_FilePass(blob, length, opts) {
	var o = { Type: blob.read_shift(2) }; /* wEncryptionType */
	if(o.Type) parse_FilePassHeader(blob, length-2, o);
	else parse_XORObfuscation(blob, length-2, opts, o);
	return o;
}


