Attribute VB_Name = "Consts"
Option Explicit

' Algorithm classes
Public Const ALG_CLASS_ANY          As Long = 0
Public Const ALG_CLASS_SIGNATURE    As Long = 8192
Public Const ALG_CLASS_MSG_ENCRYPT  As Long = 16384
Public Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Public Const ALG_CLASS_HASH         As Long = 32768
Public Const ALG_CLASS_KEY_EXCHANGE As Long = 40960
Public Const ALG_CLASS_ALL          As Long = 57344

' Algorithm types
Public Const ALG_TYPE_ANY            As Long = 0
Public Const ALG_TYPE_DSS            As Long = 512
Public Const ALG_TYPE_RSA            As Long = 1024
Public Const ALG_TYPE_BLOCK          As Long = 1536
Public Const ALG_TYPE_STREAM         As Long = 2048
Public Const ALG_TYPE_DH             As Long = 2560
Public Const ALG_TYPE_SECURECHANNEL  As Long = 3072
 
' Generic sub-ids
Public Const ALG_SID_ANY  As Long = 0

' Some RSA sub-ids
Public Const ALG_SID_RSA_ANY        As Long = 0
Public Const ALG_SID_RSA_PKCS       As Long = 1
Public Const ALG_SID_RSA_MSATWORK   As Long = 2
Public Const ALG_SID_RSA_ENTRUST    As Long = 3
Public Const ALG_SID_RSA_PGP        As Long = 4

' Some DSS sub-ids
'
Public Const ALG_SID_DSS_ANY     As Long = 0
Public Const ALG_SID_DSS_PKCS    As Long = 1
Public Const ALG_SID_DSS_DMS     As Long = 2

' Block cipher sub ids
' DES sub_ids
Public Const ALG_SID_DES            As Long = 1
Public Const ALG_SID_3DES           As Long = 3
Public Const ALG_SID_DESX           As Long = 4
Public Const ALG_SID_IDEA           As Long = 5
Public Const ALG_SID_CAST           As Long = 6
Public Const ALG_SID_SAFERSK64      As Long = 7
Public Const ALG_SID_SAFERSK128     As Long = 8
Public Const ALG_SID_3DES_112       As Long = 9
Public Const ALG_SID_CYLINK_MEK     As Long = 12
Public Const ALG_SID_RC5            As Long = 13
Public Const ALG_SID_AES_128        As Long = 14
Public Const ALG_SID_AES_192        As Long = 15
Public Const ALG_SID_AES_256        As Long = 16
Public Const ALG_SID_AES            As Long = 17

' Fortezza sub-ids
Public Const ALG_SID_SKIPJACK   As Long = 10
Public Const ALG_SID_TEK        As Long = 11

' KP_MODE
Public Const CRYPT_MODE_CBCI    As Long = 6
Public Const CRYPT_MODE_CFBP    As Long = 7
Public Const CRYPT_MODE_OFBP    As Long = 8
Public Const CRYPT_MODE_CBCOFM  As Long = 9
Public Const CRYPT_MODE_CBCOFMI As Long = 10

' RC2 sub-ids
Public Const ALG_SID_RC2 As Long = 2

' Stream cipher sub-ids
Public Const ALG_SID_RC4    As Long = 1
Public Const ALG_SID_SEAL   As Long = 2

' Diffie-Hellman sub-ids
Public Const ALG_SID_DH_SANDF       As Long = 1
Public Const ALG_SID_DH_EPHEM       As Long = 2
Public Const ALG_SID_AGREED_KEY_ANY As Long = 3
Public Const ALG_SID_KEA            As Long = 4

' Hash sub ids
Public Const ALG_SID_MD2                As Long = 1
Public Const ALG_SID_MD4                As Long = 2
Public Const ALG_SID_MD5                As Long = 3
Public Const ALG_SID_SHA                As Long = 4
Public Const ALG_SID_SHA1               As Long = 4
Public Const ALG_SID_MAC                As Long = 5
Public Const ALG_SID_RIPEMD             As Long = 6
Public Const ALG_SID_RIPEMD160          As Long = 7
Public Const ALG_SID_SSL3SHAMD5         As Long = 8
Public Const ALG_SID_HMAC               As Long = 9
Public Const ALG_SID_TLS1PRF            As Long = 10
Public Const ALG_SID_HASH_REPLACE_OWF   As Long = 11
Public Const ALG_SID_SHA_256            As Long = 12
Public Const ALG_SID_SHA_384            As Long = 13
Public Const ALG_SID_SHA_512            As Long = 14

' secure channel sub ids
Public Const ALG_SID_SSL3_MASTER            As Long = 1
Public Const ALG_SID_SCHANNEL_MASTER_HASH   As Long = 2
Public Const ALG_SID_SCHANNEL_MAC_KEY       As Long = 3
Public Const ALG_SID_PCT1_MASTER            As Long = 4
Public Const ALG_SID_SSL2_MASTER            As Long = 5
Public Const ALG_SID_TLS1_MASTER            As Long = 6
Public Const ALG_SID_SCHANNEL_ENC_KEY       As Long = 7

' algorithm identifier definitions
Public Const CALG_MD2  As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD2)
Public Const CALG_MD4  As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD4)
Public Const CALG_MD5  As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Public Const CALG_SHA  As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)
Public Const CALG_SHA1  As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1)
Public Const CALG_MAC  As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MAC)
Public Const CALG_RSA_SIGN  As Long = (ALG_CLASS_SIGNATURE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
Public Const CALG_DSS_SIGN  As Long = (ALG_CLASS_SIGNATURE Or ALG_TYPE_DSS Or ALG_SID_DSS_ANY)
Public Const CALG_NO_SIGN  As Long = (ALG_CLASS_SIGNATURE Or ALG_TYPE_ANY Or ALG_SID_ANY)
Public Const CALG_RSA_KEYX  As Long = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_RSA Or ALG_SID_RSA_ANY)
Public Const CALG_DES  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DES)
Public Const CALG_3DES_112  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES_112)
Public Const CALG_3DES  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_3DES)
Public Const CALG_DESX  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_DESX)
Public Const CALG_RC2  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2)
Public Const CALG_RC4  As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Public Const CALG_SEAL  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_SEAL)
Public Const CALG_DH_SF  As Long = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_DH Or ALG_SID_DH_SANDF)
Public Const CALG_DH_EPHEM  As Long = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_DH Or ALG_SID_DH_EPHEM)
Public Const CALG_AGREEDKEY_ANY  As Long = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_DH Or ALG_SID_AGREED_KEY_ANY)
Public Const CALG_KEA_KEYX  As Long = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_DH Or ALG_SID_KEA)
Public Const CALG_HUGHES_MD5  As Long = (ALG_CLASS_KEY_EXCHANGE Or ALG_TYPE_ANY Or ALG_SID_MD5)
Public Const CALG_SKIPJACK  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_SKIPJACK)
Public Const CALG_TEK  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_TEK)
Public Const CALG_CYLINK_MEK  As Long = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_CYLINK_MEK)
Public Const CALG_SSL3_SHAMD5  As Long = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SSL3SHAMD5)
Public Const CALG_SSL3_MASTER  As Long = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SSL3_MASTER)
Public Const CALG_SCHANNEL_MASTER_HASH As Long = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_MASTER_HASH)
Public Const CALG_SCHANNEL_MAC_KEY  As Long = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_MAC_KEY)
Public Const CALG_SCHANNEL_ENC_KEY  As Long = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SCHANNEL_ENC_KEY)
Public Const CALG_PCT1_MASTER  As Long = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_PCT1_MASTER)
Public Const CALG_SSL2_MASTER  As Long = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_SSL2_MASTER)
Public Const CALG_TLS1_MASTER = (ALG_CLASS_MSG_ENCRYPT Or ALG_TYPE_SECURECHANNEL Or ALG_SID_TLS1_MASTER)
Public Const CALG_RC5 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC5)
Public Const CALG_HMAC = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_HMAC)
Public Const CALG_TLS1PRF = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_TLS1PRF)
Public Const CALG_HASH_REPLACE_OWF = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_HASH_REPLACE_OWF)
Public Const CALG_AES_128 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_128)
Public Const CALG_AES_192 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_192)
Public Const CALG_AES_256 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_256)
Public Const CALG_AES = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES)
Public Const CALG_SHA_256 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_256)
Public Const CALG_SHA_384 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_384)
Public Const CALG_SHA_512 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA_512)

' digital signature
Public Const PROV_RSA_FULL          As Long = 1
Public Const PROV_RSA_SIG           As Long = 2
Public Const PROV_DSS               As Long = 3
Public Const PROV_FORTEZZA          As Long = 4
Public Const PROV_MS_EXCHANGE       As Long = 5
Public Const PROV_SSL               As Long = 6
Public Const PROV_RSA_SCHANNEL      As Long = 12
Public Const PROV_DSS_DH            As Long = 13
Public Const PROV_EC_ECDSA_SIG      As Long = 14
Public Const PROV_EC_ECNRA_SIG      As Long = 15
Public Const PROV_EC_ECDSA_FULL     As Long = 16
Public Const PROV_EC_ECNRA_FULL     As Long = 17
Public Const PROV_DH_SCHANNEL       As Long = 18
Public Const PROV_SPYRUS_LYNKS      As Long = 20
Public Const PROV_RNG               As Long = 21
Public Const PROV_INTEL_SEC         As Long = 22
Public Const PROV_REPLACE_OWF       As Long = 23
Public Const PROV_RSA_AES           As Long = 24

' crypto providers
Public Const MS_DEF_PROV$ = "Microsoft Base Cryptographic Provider v1.0"
Public Const MS_ENHANCED_PROV$ = "Microsoft Enhanced Cryptographic Provider v1.0"
Public Const MS_STRONG_PROV$ = "Microsoft Strong Cryptographic Provider"
Public Const MS_DEF_RSA_SIG_PROV$ = "Microsoft RSA Signature Cryptographic Provider"
Public Const MS_DEF_RSA_SCHANNEL_PROV$ = "Microsoft RSA SChannel Cryptographic Provider"
Public Const MS_DEF_DSS_PROV$ = "Microsoft Base DSS Cryptographic Provider"
Public Const MS_DEF_DSS_DH_PROV$ = "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider"
Public Const MS_ENH_DSS_DH_PROV$ = "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider"
Public Const MS_DEF_DH_SCHANNEL_PROV$ = "Microsoft DH SChannel Cryptographic Provider"
Public Const MS_SCARD_PROV$ = "Microsoft Base Smart Card Crypto Provider"
Public Const MS_ENH_RSA_AES_PROV$ = "Microsoft Enhanced RSA and AES Cryptographic Provider"

