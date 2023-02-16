'
' Copyright (c) Microsoft Corporation.  All rights reserved.
'
' VBScript Source File 
'
' Script Name: winrm.vbs
'

Option Explicit

'''''''''''''''''''''
' Error codes
private const ERR_OK              = 0
private const ERR_GENERAL_FAILURE = 1

'''''''''''''''''''''
' Messages
private const L_ONLYCSCRIPT_Message     = "Can be executed only by cscript.exe."
private const L_UNKOPNM_Message         = "Unknown operation name: "
private const L_OP_Message              = "Operation - "
private const L_NOFILE_Message          = "File does not exist: "
private const L_PARZERO_Message         = "Parameter is zero length #"
private const L_INVOPT_Message          = "Invalid option: "
private const L_UNKOPT_Message          = "Unknown option: "
private const L_BADMATCNT1_Message      = "Unexpected match count - one match is expected: "
private const L_OPTNOTUNQ_Message       = "Option is not unique: "
private const L_URIZERO_Message         = "URI is 0 length"    
private const L_URIZEROTOK_Message      = "Invalid URI, token is 0 length"    
private const L_INVWMIURI1_Message      = "Invalid WMI V1 resource URI - no '/' found  (at least 2 expected)"
private const L_INVWMIURI2_Message      = "Invalid WMI V1 resource URI - only one '/' found (at least 2 expected)"
private const L_NOLASTTOK_Message       = "Invalid URI - cannot locate last token for root node name"
private const L_ONEHASHONLY_Message     = "One hash only expected, found: "
private const L_HASHLAST_Message        = "Invalid hash or hash is not the last item on the commandline."
private const L_INVHASH1_Message        = "Invalid hash format - validation pattern failed."
private const L_INVHASH2_Message        = "Invalid hash format - hash has content but it doesn't match."
private const L_INVHASH3_Message        = "Invalid submatch number - 3 submatches expected."
private const L_PSWD_Message            = "Type password and press the Enter key:"
private const L_ARGNOVAL_Message        = "Argument's value is not provided: "   
private const L_XMLERROR_Message        = "Unable to parse XML: "         
private const L_XSLERROR_Message        = "Unable to parse XSL file. Either it is inaccessible or invalid: "         
private const L_MSXML3MISSING_Message   = "Unable to load MSXML3, required by -format option and for put using ""@{...}"""
private const L_FORMATLERROR_Message    = "Invalid option for -format: "
private const L_PUT_PARAM_NOMATCH_Message = "Parameter name does not match any properties on resource: "
private const L_PUT_PARAM_MULTIMATCH_Message = "Parameter matches more than one property on resource: "
private const L_PUT_PARAM_NOARRAY_Message = "Multiple matching parameter names not allowed in @{...}: "
private const L_PUT_PARAM_NOTATTR_Message = "Parameter matches a non-text property on resource: "
private const L_PUT_PARAM_EMPTY_Message = "Parameter set is empty."

'''''''''''''''''''''
' HELP
private const L_HELP_000_0_Message = "Windows Remote Management Command Line Tool"
private const L_HELP_002_0_Message = ""
private const L_HELP_003_0_Message = "USAGE"
private const L_HELP_004_0_Message = "====="
private const L_HELP_005_0_Message = ""
private const L_HELP_006_0_Message = "(ALL UPPER-CASE = VARIABLE)"
private const L_HELP_007_0_Message = ""
private const X_HELP_008_0_Message = "winrm { get | put | enumerate | create | delete } RESOURCE_URI "
private const L_HELP_009_0_Message = " [-OPTION:VALUE [-OPTION:VALUE] ...]  [@{KEY=VALUE[;KEY=VALUE]...}]"
private const L_HELP_010_0_Message = ""
private const L_HELP_011_0_Message = "winrm invoke ACTION RESOURCE_URI [-OPTION:VALUE [-OPTION:VALUE] ...]"
private const L_HELP_012_0_Message = " [@{KEY=VALUE[;KEY=VALUE]...}]"
private const L_HELP_015_0_Message = ""
private const X_HELP_016_0_Message = "winrm -?"
private const L_HELP_017_0_Message = ""
private const L_HELP_023_0_Message = ""
private const L_HELP_024_0_Message = "OPERATIONS"
private const L_HELP_025_0_Message = "=========="
private const L_HELP_026_0_Message = ""
private const X_HELP_027_0_Message = "get"
private const L_HELP_028_0_Message = "---"
private const L_HELP_029_0_Message = "Retrieves instances of RESOURCE_URI using specified options and key-value"
private const L_HELP_030_0_Message = "pairs."
private const L_HELP_031_0_Message = ""
private const X_HELP_032_0_Message = "put"
private const L_HELP_033_0_Message = "---"
private const L_HELP_034_0_Message = "Modifies settings in RESOURCE_URI using specified options and key-value pairs."
private const L_HELP_035_0_Message = ""
private const X_HELP_036_0_Message = "enumerate"
private const L_HELP_037_0_Message = "---------"
private const L_HELP_038_0_Message = "Lists instances of RESOURCE_URI."
private const L_HELP_039_0_Message = ""
private const X_HELP_040_0_Message = "create"
private const L_HELP_041_0_Message = "------"
private const L_HELP_042_0_Message = "Spawns an instance of RESOURCE_URI using specified key-value pairs."
private const L_HELP_043_0_Message = ""
private const X_HELP_044_0_Message = "delete"
private const L_HELP_045_0_Message = "------"
private const L_HELP_046_0_Message = "Removes an instance of RESOURCE_URI."
private const L_HELP_047_0_Message = ""
private const X_HELP_048_0_Message = "invoke"
private const L_HELP_049_0_Message = "------"
private const L_HELP_050_0_Message = "Executes method specified by ACTION on target object specified by "
private const L_HELP_051_0_Message = "RESOURCE_URI with parameters specified by key-value pairs."
private const L_HELP_052_0_Message = ""
private const X_HELP_065_0_Message = "-? (or no parameters)"
private const L_HELP_066_0_Message = "---------------------"
private const L_HELP_067_0_Message = "Displays this syntax listing."
private const L_HELP_068_0_Message = ""
private const L_HELP_069_0_Message = ""
private const L_HELP_070_0_Message = "ARGUMENTS"
private const L_HELP_071_0_Message = "========="
private const L_HELP_072_0_Message = ""
private const L_HELP_073_0_Message = "RESOURCE_URI"
private const L_HELP_074_0_Message = "------------"
private const L_HELP_075_0_Message = "Universal Resource Identifier specifying resource to be used for operation."
private const L_HELP_076_0_Message = "Selectors and values are passed after the URI in the form:"
private const L_HELP_077_0_Message = "?name=value[+name=value]..."
private const L_HELP_078_0_Message = "Example:"
private const X_HELP_079_0_Message = "wsman:microsoft.com/wsman/2005/06/config/Listener?IP=1.2.3.4+Port=443"
private const L_HELP_080_0_Message = "Note: Some parts of RESOURCE_URI are case-sensitive. When using create or"
private const L_HELP_081_0_Message = "invoke, the last part of the resource URI must match case-wise the top-level"
private const L_HELP_082_0_Message = "element of the expected XML."
private const L_HELP_083_0_Message = ""
private const L_HELP_084_0_Message = "Aliasing allows shortcuts to be used in place of full URIs."
private const L_HELP_085_0_Message = "Available aliases and the URIs they substitute for are:"
private const X_HELP_086_0_Message = "wmi = http://schemas.microsoft.com/wsman/2005/06/wmi"
private const X_HELP_087_0_Message = "wsman = wsman:microsoft.com/wsman/2005/06/"
private const X_HELP_088_0_Message = "cimv2.9 = http://schemas.dmtf.org/wsman/2005/06/cimv2.9"
private const X_HELP_089_0_Message = "cimv2 = http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2"
private const L_HELP_090_0_Message = "Example:"
private const L_HELP_091_0_Message = "Using full URI:"
private const X_HELP_092_0_Message = "winrm get http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Service?Name=WSMan"
private const L_HELP_093_0_Message = "Using alias:"
private const X_HELP_094_0_Message = "winrm get cimv2/Win32_Service?Name=WSMan"
private const L_HELP_095_0_Message = ""
private const L_HELP_096_0_Message = ""
private const L_HELP_097_0_Message = "ACTION"
private const L_HELP_098_0_Message = "------"
private const L_HELP_099_0_Message = "Method of RESOURCE_URI to be executed by invoke operation."
private const L_HELP_100_0_Message = ""
private const L_HELP_101_0_Message = ""
private const L_HELP_102_0_Message = "OPTIONS"
private const L_HELP_103_0_Message = "======="
private const L_HELP_104_0_Message = ""
private const X_HELP_105_0_Message = "-machine:VALUE"
private const L_HELP_106_0_Message = "--------------"
private const L_HELP_107_0_Message = "Specifies identifier of remote endpoint/machine.  Can be in the form of a DNS "
private const L_HELP_108_0_Message = "name, NetBIOS name, or IP address. When using HTTPS, the machine name must"
private const L_HELP_109_0_Message = "match the server's certificate common name (CN) unless -skipCNcheck is used."
private const L_HELP_110_0_Message = "Examples:"
private const X_HELP_111_0_Message = "-machine:srv.corp.com"
private const X_HELP_112_0_Message = "-machine:machine1"
private const X_HELP_113_0_Message = "-machine:127.0.0.1"
private const L_HELP_114_0_Message = ""
private const X_HELP_115_0_Message = "-port:VALUE"
private const L_HELP_116_0_Message = "-----------"
private const L_HELP_117_0_Message = "Specifies communication port on the target system to perform operation on."
private const L_HELP_118_0_Message = "If not specified, previously configured default port is used."
private const L_HELP_119_0_Message = "Example:"
private const X_HELP_120_0_Message = "-port:9999"
private const L_HELP_121_0_Message = ""
private const X_HELP_122_0_Message = "-transport:VALUE"
private const L_HELP_123_0_Message = "----------------"
private const L_HELP_124_0_Message = "Specifies communication protocol to use. Available protocols are http and https."
private const L_HELP_125_0_Message = "Default is https."
private const L_HELP_126_0_Message = "Example:"
private const L_HELP_127_0_Message = "-transport:http"
private const L_HELP_128_0_Message = ""
private const X_HELP_129_0_Message = "-prefix:VALUE"
private const L_HELP_130_0_Message = "-------------"
private const L_HELP_131_0_Message = "HTTP prefix to use on the URL used to talk to a remote machine."
private const L_HELP_132_0_Message = "Default is ""wsman""."
private const L_HELP_133_0_Message = "Examples:"
private const X_HELP_134_0_Message = "-prefix:wsman"
private const L_HELP_135_0_Message = "  Construct URL in form http://MACHINE/wsman, using wsman to route request."
private const L_HELP_136_0_Message = ""
private const X_HELP_137_0_Message = "-timeout:MS"
private const L_HELP_138_0_Message = "-----------"
private const L_HELP_139_0_Message = "Timeout in milliseconds. Limits duration of corresponding operation."
private const L_HELP_140_0_Message = "Default timeout can be configured by:"
private const L_HELP_141_0_Message = "winrm put wsman:microsoft.com/wsman/2005/06/Config @{MaxTimeoutms=""XXXXXX""}"
private const L_HELP_142_0_Message = "Where XXXXXX is an integer indicating milliseconds."
private const L_HELP_143_0_Message = ""
private const X_HELP_144_0_Message = "-auth:VALUE"
private const L_HELP_145_0_Message = "-----------"
private const L_HELP_146_0_Message = "Specifies authorization mechanism used when communicating with remote"
private const L_HELP_147_0_Message = "machine (see -machine option)."
private const L_HELP_148_0_Message = "Possible options are ""Basic"", ""Digest"" and ""Negotiate""."
private const L_HELP_149_0_Message = "The default is ""Negotiate""."
private const L_HELP_150_0_Message = "Examples:"
private const X_HELP_151_0_Message = "-auth:Basic"
private const X_HELP_152_0_Message = "-auth:Digest"
private const X_HELP_153_0_Message = "-auth:Negotiate"
private const L_HELP_155_0_Message = ""
private const X_HELP_156_0_Message = "-username:USERNAME"
private const L_HELP_157_0_Message = "Specifies username on remote machine (see -machine option)."
private const L_HELP_158_0_Message = "User must be member of local Administrators group on remote machine."
private const L_HELP_159_0_Message = "Cannot be used on local machine."
private const L_HELP_160_0_Message = "If the user account is a local account on the remote machine, "
private const L_HELP_161_0_Message = "the syntax should be in the form -username:USERNAME"
private const L_HELP_162_0_Message = "If the username is a domain account, the syntax should be in the form "
private const X_HELP_163_0_Message = "-username:DOMAIN\USERNAME"
private const L_HELP_164_0_Message = "If -username option is not used, then current logged-on user's credentials are "
private const L_HELP_165_0_Message = "used. In this case, logged-on account must be member of local "
private const L_HELP_166_0_Message = "Administrators group on remote machine."
private const L_HELP_167_0_Message = ""
private const X_HELP_188_0_Message = "-password:PASSWORD"
private const L_HELP_189_0_Message = "------------------"
private const L_HELP_190_0_Message = "Specifies password on command line to override interactive prompt."
private const L_HELP_191_0_Message = "Applies only if -username:USERNAME option is used."
private const L_HELP_192_0_Message = ""
private const X_HELP_193_0_Message = "-skipCAcheck"
private const L_HELP_194_0_Message = "------------"
private const L_HELP_195_0_Message = "Specifies that certificate issuer need not be a trusted root authority. "
private const L_HELP_196_0_Message = "Used only in remote operations using HTTPS"
private const L_HELP_197_0_Message = "(see -machine option)."
private const L_HELP_198_0_Message = "This option should be used only for trusted machines."
private const L_HELP_198_1_Message = ""
private const X_HELP_198_2_Message = "-dialect:<value>"
private const L_HELP_198_3_Message = "-----------------"
private const L_HELP_198_4_Message = "Dialect of the filter expression for enumeration."
private const L_HELP_198_5_Message = ""
private const X_HELP_198_6_Message = "-filter:<value>"
private const L_HELP_198_7_Message = "-----------------"
private const L_HELP_198_8_Message = "Filter expression for enumeration."
private const L_HELP_199_0_Message = ""
private const X_HELP_200_0_Message = "-skipCNcheck"
private const L_HELP_201_0_Message = "------------"
private const L_HELP_202_0_Message = "Specifies that certificate common name (CN) of the server need not match the"
private const L_HELP_203_0_Message = "hostname of the server. Used only in remote operations using HTTPS"
private const L_HELP_204_0_Message = "(see -machine option)."
private const L_HELP_205_0_Message = "This option should be used only for trusted machines."
private const L_HELP_206_0_Message = ""
private const X_HELP_206_1_Message = "-SPNPort"
private const L_HELP_206_2_Message = "------------"
private const L_HELP_206_3_Message = "Appends port number to the service principal name of the remote server."
private const L_HELP_206_4_Message = "Service principal name is used when Negotiate authentication mechanism is in use."
private const L_HELP_206_5_Message = ""
private const X_HELP_207_0_Message = "-encoding:VALUE"
private const L_HELP_208_0_Message = "---------------"
private const L_HELP_209_0_Message = "Specifies encoding type when talking to remote machine (see -machine "
private const L_HELP_210_0_Message = "option). Possible options are utf-8 and utf-16."
private const L_HELP_211_0_Message = "Default is utf-16."
private const L_HELP_212_0_Message = "Examples:"
private const X_HELP_213_0_Message = "-encoding:utf-8"
private const X_HELP_214_0_Message = "-encoding:utf-16"
private const L_HELP_215_0_Message = ""
private const X_HELP_216_0_Message = "-file:VALUE"
private const L_HELP_217_0_Message = "-----------"
private const L_HELP_218_0_Message = "Specifies name of file used as input."
private const L_HELP_219_0_Message = "VALUE can be absolute path, relative path, or filename without path."
private const L_HELP_220_0_Message = "Names or paths that include spaces must be enclosed in quotation marks."
private const L_HELP_221_0_Message = ""
private const X_HELP_222_0_Message = "-strict"
private const L_HELP_223_0_Message = "-------"
private const L_HELP_224_0_Message = "Reject unknown parameters. This ensures that the script does not ignore"
private const L_HELP_225_0_Message = "parameters that it does not understand."
private const L_HELP_226_0_Message = ""
private const X_HELP_227_0_Message = "-format:#FORMAT"
private const L_HELP_228_0_Message = "---------------"
private const L_HELP_229_0_Message = "Specifies format of output."
private const L_HELP_230_0_Message = "FORMAT can be ""xml"", ""pretty"" (better formatted XML), or ""text""."
private const L_HELP_231_0_Message = "Examples:"
private const X_HELP_232_0_Message = "-format:#xml"
private const X_HELP_233_0_Message = "-format:#pretty"
private const X_HELP_234_0_Message = "-format:#text"
private const L_HELP_235_0_Message = ""
private const L_HELP_236_0_Message = ""
private const L_HELP_237_0_Message = "@{KEY=VALUE} HASH TABLE"
private const L_HELP_238_0_Message = "========================"
private const L_HELP_239_0_Message = "Key/value pairs list for operations that accept input."
private const L_HELP_240_0_Message = "Must be last item on command line."
private const L_HELP_241_0_Message = "Keys are not unique. Values must be within quotation marks."
private const L_HELP_242_0_Message = "Use either @{KEY=VALUE} or input from an XML file using -file, but not both."
private const L_HELP_243_0_Message = ""
private const L_HELP_244_0_Message = "@{KEY=VALUE[;KEY=VALUE]}"
private const L_HELP_245_0_Message = "------------------------"
private const L_HELP_246_0_Message = "Supply property names and values."
private const L_HELP_247_0_Message = ""
private const L_HELP_248_0_Message = "@{KEY=$null}"
private const L_HELP_249_0_Message = "------------"
private const L_HELP_250_0_Message = "$null is a special value."
private const L_HELP_251_0_Message = ""
private const L_HELP_252_0_Message = "Examples:"
private const L_HELP_253_0_Message = ""
private const L_HELP_254_0_Message = "@{}"
private const L_HELP_255_0_Message = "No input. Required for invoke operations that take no input."
private const L_HELP_256_0_Message = ""
private const L_HELP_257_0_Message = "@{key1=""value1"";key2=""value2""}"
private const L_HELP_258_0_Message = ""
private const L_HELP_259_0_Message = "@{key1=$null;key2=""value2""}"
private const L_HELP_260_0_Message = ""
private const L_HELP_261_0_Message = ""
private const L_HELP_262_0_Message = "Examples"
private const L_HELP_263_0_Message = "========"
private const L_HELP_264_0_Message = ""
private const X_HELP_265_0_Message = "get"
private const L_HELP_266_0_Message = "---"
private const L_HELP_267_0_Message = ""
private const L_HELP_268_0_Message = "Retrieve current configuration in XML format:"
private const X_HELP_269_0_Message = "winrm get wsman:microsoft.com/wsman/2005/06/config"
private const L_HELP_270_0_Message = ""
private const L_HELP_271_0_Message = "Retrieve spooler instance of Win32_Service class:"
private const X_HELP_272_0_Message = "winrm get http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Service?Name=spooler"
private const L_HELP_273_0_Message = ""
private const L_HELP_274_0_Message = "Check whether the Windows Remote Management service is installed and running on a remote"
private const L_HELP_275_0_Message = "machine:"
private const X_HELP_276_0_Message = "winrm get wsman:system/2005/06/this -machine:host.acme.com"
private const L_HELP_277_0_Message = ""
private const X_HELP_278_0_Message = "put"
private const L_HELP_279_0_Message = "---"
private const L_HELP_280_0_Message = ""
private const L_HELP_281_0_Message = "Modify a configuration property of Windows Remote Management:"
private const X_HELP_282_0_Message = "winrm put wsman:microsoft.com/wsman/2005/06/Config @{MaxEnvelopeSizekb=""100""}"
private const L_HELP_283_0_Message = ""
private const L_HELP_284_0_Message = "Enable Listener on a remote machine:"
private const X_HELP_285_0_Message = "winrm put wsman:microsoft.com/wsman/2005/06/config/Listener?IP=1.2.3.4+Port=443 @{Enabled=""True""}"
private const L_HELP_286_0_Message = ""
private const L_HELP_287_0_Message = "Disable Listener on a remote machine:"
private const X_HELP_288_0_Message = "winrm put wsman:microsoft.com/wsman/2005/06/config/Listener?IP=192.168.2.1+Port=443 @{Enabled=""False""}"
private const L_HELP_289_0_Message = ""
private const L_HELP_290_0_Message = "Set Enabled property of Listener by reading from an XML file:"
private const L_HELP_291_0_Message = "winrm put wsman:microsoft.com/wsman/2005/06/config/Listener?IP=1.2.3.4+Port=443 -file:config.xml"
private const L_HELP_292_0_Message = "Where config.xml:"
private const X_HELP_293_0_Message = "<cfg:Listener xmlns:cfg=""wsman:microsoft.com/wsman/2005/06/config/listener.xsd"">"
private const X_HELP_294_0_Message = "   <cfg:Enabled>False</cfg:Enabled>"
private const X_HELP_295_0_Message = "</cfg:Listener>"
private const L_HELP_296_0_Message = ""
private const X_HELP_297_0_Message = "enumerate"
private const L_HELP_298_0_Message = "---------"
private const L_HELP_299_0_Message = ""
private const L_HELP_300_0_Message = "List all instances of Listener:"
private const X_HELP_301_0_Message = "winrm enumerate wsman:microsoft.com/wsman/2005/06/config/Listener"
private const L_HELP_302_0_Message = ""
private const L_HELP_303_0_Message = "List all instances of Win32_Service class:"
private const L_HELP_304_0_Message = "winrm enumerate http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Service"
private const L_HELP_305_0_Message = ""
private const X_HELP_306_0_Message = "create"
private const L_HELP_307_0_Message = "------"
private const L_HELP_308_0_Message = ""
private const L_HELP_309_0_Message = "Create instance of Listener:"
private const X_HELP_310_0_Message = "winrm create wsman:microsoft.com/wsman/2005/06/config/Listener?IP=192.168.2.1+Port=443  @{Hostname=""host"";CertificateThumbprint=""XXXXXXXXXX""}"
private const L_HELP_311_0_Message = "Note: XXXXXXXXXX represents a 40-digit hex string."
private const L_HELP_312_0_Message = ""
private const L_HELP_313_0_Message = "Following example uses IPv6 address to create listener."
private const L_HELP_314_0_Message = "winrm create wsman:microsoft.com/wsman/2005/06/config/Listener?IP=3ffe:8311:ffff:f2c1:211:11ff:fe81:5e61+Port=443  @{Hostname=""host"";CertificateThumbprint=""XXXXXXXXXX""}"
private const L_HELP_315_0_Message = ""
private const L_HELP_316_0_Message = ""
private const X_HELP_317_0_Message = "delete"
private const L_HELP_318_0_Message = "------"
private const L_HELP_319_0_Message = ""
private const L_HELP_320_0_Message = "Delete an instance of Listener:"
private const X_HELP_321_0_Message = "winrm delete wsman:microsoft.com/wsman/2005/06/config/Listener?IP=192.168.2.1+Port=444"
private const L_HELP_322_0_Message = ""
private const X_HELP_323_0_Message = "invoke"
private const L_HELP_324_0_Message = "------"
private const L_HELP_325_0_Message = ""
private const L_HELP_326_0_Message = "Call StartService method on spooler instance of Win32_Service class:"
private const X_HELP_327_0_Message = "winrm invoke StartService http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Service?Name=spooler @{}"
private const L_HELP_328_0_Message = ""
private const L_HELP_329_0_Message = "Call StopService method on spooler instance of of Win32_Service class using"
private const L_HELP_330_0_Message = "XML file:"
private const X_HELP_331_0_Message = "winrm invoke StopService http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Service?Name=spooler -file:input.xml"
private const L_HELP_332_0_Message = "Where input.xml:"
private const X_HELP_333_0_Message = "<p:StopService_INPUT xmlns:p=""http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Service""/>"
private const L_HELP_334_0_Message = ""
private const L_HELP_335_0_Message = "Call Create method of Win32_Process class with specified parameters:"
private const X_HELP_336_0_Message = "winrm invoke Create http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Process @{CommandLine=""notepad.exe"";CurrentDirectory=""C:\""}"
private const L_HELP_337_0_Message = ""
private const L_HELP_338_0_Message = "Retrieve the thumbprint of a certificate:"
private const X_HELP_339_0_Message = "winrm invoke Retrieve wsman:Microsoft.com/wsman/2005/06/config/Certificate @{Hostname=""host.acme.com""}"
private const L_HELP_340_0_Message = ""
private const L_HELP_341_0_Message = "Call Create method of Win32_Process class with parameters specified in XML"
private const L_HELP_342_0_Message = "file:"
private const X_HELP_343_0_Message = "winrm invoke Create http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Process -file:input.xml"
private const L_HELP_344_0_Message = "Where input.xml:"
private const X_HELP_345_0_Message = "<p:Create_INPUT xmlns:p=""http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/Win32_Process"">"
private const X_HELP_346_0_Message = "  <p:CommandLine>notepad.exe</p:CommandLine>"
private const X_HELP_347_0_Message = "  <p:CurrentDirectory>C:\</p:CurrentDirectory>"
private const X_HELP_348_0_Message = "</p:Create_INPUT>"




'''''''''''''''''''''
' Operations
private const OP_GET  = "get"
private const OP_PUT  = "put"
private const OP_CRE  = "create"
private const OP_DEL  = "delete"
private const OP_ENU  = "enumerate"
private const OP_INV  = "invoke"

'''''''''''''''''''''
' Named parameters (key names of key/value pairs)
private const NPARA_USERNAME = "username"
private const NPARA_PASSWORD = "password"
private const NPARA_DIALECT  = "dialect"
private const NPARA_FILE     = "file"
private const NPARA_FILTER   = "filter"
private const NPARA_HELP     = "?"
private const NPARA_PORT     = "port"
private const NPARA_PREFIX   = "prefix"
private const NPARA_STRICT   = "strict"
private const NPARA_MACHINE   = "machine"
private const NPARA_TRANSPORT = "transport"
private const NPARA_NOCACHK  = "skipcacheck"
private const NPARA_NOCNCHK  = "skipcncheck"
private const NPARA_SPNPORT  = "spnport"
private const NPARA_TIMEOUT  = "timeout"
private const NPARA_TRACE    = "trace"
private const NPARA_AUTH     = "auth"
private const NPARA_ENCODING = "encoding"
private const NPARA_FORMAT   = "format"
' Literal values in key/value pairs
private const VAL_BASIC      = "basic"
private const VAL_DIGEST     = "digest"
private const VAL_NEGOTIATE  = "negotiate"
' Output formatting flags
private const VAL_FORMAT_XML         = "#xml"
private const VAL_FORMAT_PRETTY      = "#pretty"
private const VAL_FORMAT_PRETTY_XSLT = "WsmPty.xsl"
private const VAL_FORMAT_TEXT        = "#text"
private const VAL_FORMAT_TEXT_XSLT   = "WsmTxt.xsl"

private const VAL_ALIAS_WMI_NAME =          "wmi"
private const VAL_ALIAS_WMI_VALUE =         "http://schemas.microsoft.com/wsman/2005/06/wmi/"
private const VAL_ALIAS_WSMAN_NAME =        "wsman"
private const VAL_ALIAS_WSMAN_VALUE =       "wsman:microsoft.com/wsman/2005/06/"
private const VAL_ALIAS_CIMV2_9_NAME =      "cimv2.9"
private const VAL_ALIAS_CIMV2_9_VALUE =     "http://schemas.dmtf.org/wsman/2005/06/cimv2.9/"
private const VAL_ALIAS_CIMV2_NAME =        "cimv2"
private const VAL_ALIAS_CIMV2_VALUE =       "http://schemas.microsoft.com/wsman/2005/06/wmi/root/cimv2/"

'''''''''''''''''''''
' Patterns
private const PTRN_IPV6_1 = "([A-Fa-f0-9]{1,4}:){6}:[A-Fa-f0-9]{1,4}"
private const PTRN_IPV6_2 = "([A-Fa-f0-9]{1,4}:){7}[A-Fa-f0-9]{1,4}"
private const PTRN_IPV6_3 = "[A-Fa-f0-9]{1,4}::([A-Fa-f0-9]{1,4}:){0,5}[A-Fa-f0-9]{1,4}"
private const PTRN_IPV6_4 = "([A-Fa-f0-9]{1,4}:){2}:([A-Fa-f0-9]{1,4}:){0,4}[A-Fa-f0-9]{1,4}"
private const PTRN_IPV6_5 = "([A-Fa-f0-9]{1,4}:){3}:([A-Fa-f0-9]{1,4}:){0,3}[A-Fa-f0-9]{1,4}"
private const PTRN_IPV6_6 = "([A-Fa-f0-9]{1,4}:){4}:([A-Fa-f0-9]{1,4}:){0,2}[A-Fa-f0-9]{1,4}"
private const PTRN_IPV6_7 = "([A-Fa-f0-9]{1,4}:){5}:([A-Fa-f0-9]{1,4}:){0,1}[A-Fa-f0-9]{1,4}"
private const PTRN_IPV6_S = ":"

private const PTRN_OPT = "^-([a-z]+):(.*)"
private const PTRN_HASH_PART = "@\{"
private const PTRN_HASH = "@\{[^{}]*\}"
private const PTRN_HASH_0 = "@\{\s*\}"
private const PTRN_HASH_LAST = "@\{.*\}\s*(/trace)?\s*$"
private const PTRN_HASH_TOK = "\s*(\w+)\s*=\s*(\$null|""([^""]*)"")\s*"
private const PTRN_URI_LAST = "([a-z_][-a-z0-9._]*)$"

dim PTRN_HASH_TOK_P
PTRN_HASH_TOK_P = "(" & PTRN_HASH_TOK & ")"

dim PTRN_HASH_VAL
PTRN_HASH_VAL = _
    "@\{(" & PTRN_HASH_TOK_P & ";)*(" & PTRN_HASH_TOK_P & ")\}"

dim PTRN_KEYS
PTRN_KEYS = "^(" & _
    NPARA_DIALECT & "|" & NPARA_FILE & "|" & _
    NPARA_FILTER  & "|" & NPARA_AUTH & "|" & NPARA_ENCODING & "|" & _
    "\" & NPARA_HELP & "|" & NPARA_PORT & "|" & NPARA_PREFIX & "|" & _
    NPARA_NOCACHK & "|" & NPARA_NOCNCHK & "|" & NPARA_SPNPORT & "|" & _
    NPARA_STRICT & "|" & NPARA_MACHINE & "|" & NPARA_TRANSPORT & "|" & _
    NPARA_TIMEOUT & "|" & NPARA_TRACE & "|" & NPARA_PASSWORD & "|" &_
    NPARA_USERNAME & "|" & NPARA_FORMAT & ")$"

dim PTRN_IPV6
PTRN_IPV6 = "^(" & _
    PTRN_IPV6_1 & ")$|^(" & PTRN_IPV6_2 & ")$|^(" & _
    PTRN_IPV6_3 & ")$|^(" & PTRN_IPV6_4 & ")$|^(" & PTRN_IPV6_5 & ")$|^(" & _
    PTRN_IPV6_6 & ")$|^(" & PTRN_IPV6_7 & ")$"

'''''''''''''''''''''    
' Session flags
private const WSManFlagEncodingUTF8             = &h00000001
private const WSManFlagCredUsernamePassword     = &h00001000
private const WSManFlagSkipCACheck              = &h00002000
private const WSManFlagSkipCNCheck              = &h00004000
private const WSManFlagUseDigest                = &h00010000
private const WSManFlagUseNegotiate             = &h00020000
private const WSManFlagUseBasic                 = &h00040000
private const WSManFlagEnableSPNServerPort      = &h00400000

'''''''''''''''''''''
' Misc
private const T_O       = &h800705B4
private const URI_IPMI  = "http://schemas.dmtf.org/wsman/2005/06/cimv2.9"
private const URI_WMIV1 = "http://schemas.microsoft.com/wsman/2005/06/wmi"
private const NS_IPMI   = "http://schemas.dmtf.org/wsman/2005/06/cimv2.9"
private const NS_WSMANL = "http://schemas.microsoft.com"
private const NS_WSMANR = "2004/07"
private const NS_XSI    = "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"""
private const ATTR_NIL  = "xsi:nil=""true"""
private const ATTR_NIL_NAME  = "xsi:nil"
private const NS_XSI_URI = "http://www.w3.org/2001/XMLSchema-instance"
private const NODE_ATTRIBUTE = 2
private const NODE_TEXT = 3


'''''''''''''''''''''
' Vars
dim wsmanObj
dim connOptionsObj
dim sessionObj
dim enumeratorObj
dim optionsDicObj
dim regexpObj

dim inputStr
dim responseStr
dim connectionStr
dim resourceUriStr
dim actionUriStr
dim rootNdNameStr
dim xmlnsStr

dim operation
dim trace
dim optStart

dim formatOption
dim formattedStr

dim stdErr
dim stdOut

set stdErr = WScript.StdErr
set stdOut = WScript.StdOut

' ------------------ Main() --------------------

ASSERTBOOL IsCScriptEnv(), L_ONLYCSCRIPT_Message

'''''''''''''''''''''    
' Parse parameters

if WScript.Arguments.Count = 0 then
    Usage(stdErr)
    WScript.Quit(ERR_GENERAL_FAILURE)
end if

if WScript.Arguments(0) = "-" & NPARA_HELP then
    Usage(stdOut)
    WScript.Quit(ERR_OK)
end if

trace = WScript.Arguments.Named.Exists(NPARA_TRACE)
TRCS WScript.Arguments.Count, "WScript.Arguments.Count"

if WScript.Arguments.Count < 2 then
    Usage(stdErr) 
    WScript.Quit(ERR_GENERAL_FAILURE)
end if

operation = GetOperation(WScript.Arguments)

optStart = 2
if operation = OP_INV then
    optStart = 3
    if WScript.Arguments.Count < 3 then
        Usage(stdErr)
        WScript.Quit(ERR_GENERAL_FAILURE)
    end if
    actionUriStr = WScript.Arguments(1)
    resourceUriStr = WScript.Arguments(2)
else
    resourceUriStr = WScript.Arguments(1)
end if

if ContainsAlias(resourceUriStr) then
    resourceUriStr = ExpandAlias(resourceUriStr)
end if

set regexpObj = new RegExp
regexpObj.Global = True
regexpObj.IgnoreCase = True

set optionsDicObj = GetOptions(regexpObj)

connectionStr = GetConnectionString(optionsDicObj, regexpObj)

set wsmanObj = CreateObject("WSMAN.Automation")

xmlnsStr = GetXmlNs(resourceUriStr)

rootNdNameStr = GetRootNodeName(operation, resourceUriStr, actionUriStr, regexpObj)

set sessionObj = CreateSession(wsmanObj, connectionStr, optionsDicObj)


if not ProcessInput(wsmanObj, regexpObj, operation, rootNdNameStr, _
            xmlnsStr, optionsDicObj, resourceUriStr,sessionObj, inputStr) then    
    WScript.Quit(ERR_GENERAL_FAILURE)
end if
            
if(optionsDicObj.Exists(NPARA_FORMAT)) then
    formatOption = optionsDicObj.Item(NPARA_FORMAT)
else
    optionsDicObj.Item(NPARA_FORMAT) = "#text"
    formatOption = "#text"
end if

'''''''''''''''''''''
' Execute operation

on error resume next
select case operation
    case OP_DEL 
        sessionObj.Delete(resourceUriStr)
    case OP_GET 
        responseStr = sessionObj.Get(resourceUriStr)
    case OP_PUT
       responseStr = sessionObj.Put(resourceUriStr, inputStr)
    case OP_CRE
        responseStr = sessionObj.Create(resourceUriStr, inputStr)
    case OP_INV
        responseStr = sessionObj.Invoke(actionUriStr, resourceUriStr, inputStr)
    case OP_ENU
         Enumerate sessionObj, optionsDicObj, resourceUriStr
    case else 
        ASSERTBOOL False, L_UNKOPNM_Message & "'" & operation & "'"
end select
if Err.Number <> 0 then
    stdErr.WriteLine sessionObj.Error
end if
ASSERTERR L_OP_Message & operation & "."
on error goto 0

if operation <> OP_ENU and operation <> OP_DEL then
    If Reformat(responseStr,formattedStr,formatOption) Then
        wscript.echo formattedStr
    End If
end if

WScript.Quit(ERR_OK)

' ------------------ Utilities --------------------

'''''''''''''''''''''    
' Reads entire text file

private function ReadFile(file)
    LTRC "ReadFile"

    dim fso, f

    set fso = CreateObject("Scripting.FileSystemObject")
    ASSERTBOOL fso.FileExists(file), L_NOFILE_Message & "'" & file & "'"
    set f = fso.OpenTextFile(file, 1, false,-2)
    ReadFile = f.ReadAll

    f.Close
    set f = Nothing
    set fso = Nothing
end function 

'''''''''''''''''''''    
' Parses switches -xxx

private function GetOptions(regexp)
    LTRC "GetOptions"

    dim i
    dim opt 
    dim matches
    dim key
    dim value
    dim strict

    strict = false
    i = optStart
    do while(i < WScript.Arguments.Count)
        TRCS WScript.Arguments(i), "(" & i & ")"    

        if WScript.Arguments(i) = "-" & NPARA_STRICT then
            LTRC "STRICT"
            strict = true
            exit do
        end if
       
        i = i + 1    
    loop
    TRCS strict, "strict"    
    
    set opt = CreateObject("Scripting.Dictionary")

    i = optStart
    do while(i < WScript.Arguments.Count)
        TRCS WScript.Arguments(i), "(" & i & ")"    

        dim l 
        l = Left(WScript.Arguments(i), 1)
        ASSERTBOOL Len(l) <> 0, L_PARZERO_Message & i
        if l = "@" then
            exit do
        end if
       
        ASSERTBOOL l = "-", L_INVOPT_Message & WScript.Arguments(i)

        regexp.Pattern = PTRN_OPT   
        regexp.IgnoreCase = TRUE
        set matches = regexp.Execute(WScript.Arguments(i))
        TRCS matches.Count, "matches.Count"    
        if matches.Count <> 0 then
            ASSERTBOOL matches.Count = 1, L_BADMATCNT1_Message & matches.Count    
            key = matches(0).SubMatches(0)
            key = LCase(key)
            value = matches(0).SubMatches(1) 
            TRCS key, "key" 
            TRCS value, "value"
            if strict then
                regexp.Pattern = PTRN_KEYS
                set matches = regexp.Execute(key)
                ASSERTBOOL matches.Count = 1, L_UNKOPT_Message & key
            end if
            ASSERTBOOL Not opt.Exists(key), L_OPTNOTUNQ_Message & key
            opt.Add key, value
        elseif LCase(WScript.Arguments(i)) = "-" & NPARA_STRICT then
            TRC "strict"
            opt.Add NPARA_STRICT, ""
        elseif LCase(WScript.Arguments(i)) = "-" & NPARA_NOCACHK then
            TRC "skipcacheck"
            opt.Add NPARA_NOCACHK, ""
        elseif LCase(WScript.Arguments(i)) = "-" & NPARA_NOCNCHK then
            TRC "skipcncheck"
            opt.Add NPARA_NOCNCHK, ""
        elseif LCase(WScript.Arguments(i)) = "-" & NPARA_SPNPORT then
            TRC "SPNPort"
            opt.Add NPARA_SPNPORT, ""
        elseif LCase(WScript.Arguments(i)) = "-" & NPARA_HELP then
            TRC "help"
            opt.Add NPARA_HELP, ""            
        else
            ASSERTBOOL false, L_INVOPT_Message & WScript.Arguments(i)
        end if
        i = i + 1    
    loop
    
    set GetOptions = opt
end function

'''''''''''''''''''''    
' Extracts operation name

private function GetOperation(arguments)
    LTRC "GetOperation"
    
    dim opr
    
    opr = LCase(arguments(0))
    select case opr
        case OP_DEL 
        case OP_GET 
        case OP_PUT 
        case OP_CRE
        case OP_INV
        case OP_ENU
        case else 
            ASSERTBOOL False, L_UNKOPNM_Message & "'" & opr & "'"
    end select
    GetOperation = opr
    TRCS GetOperation, "Operation"  
end function

private function GetConnectionString(opt, regexp)
    LTRC "GetConnectionString"
    
    dim conStr
    conStr = ""
    

    if opt.Exists(NPARA_TRANSPORT) then
        TRC "Option - " & NPARA_TRANSPORT & ",  was provided."
        conStr = opt.Item(NPARA_TRANSPORT) & "://"
    end if

    if opt.Exists(NPARA_MACHINE) then
        TRC "Option - " & NPARA_MACHINE & ",  was provided."
        ASSERTNAL(NPARA_MACHINE)
        dim tmpStr
        tmpStr = opt.Item(NPARA_MACHINE)
        if IsIPv6(tmpStr, regexp) then
            conStr = conStr  & "[" & tmpStr & "]"
    else 
            conStr = conStr & tmpStr
        end if
    else 
        'ASSERTBOOL Not opt.Exists(NPARA_PORT), "..."    
        conStr = ""
    end if
    if opt.Exists(NPARA_PORT) then
        TRC "Option - " & NPARA_PORT & ",  was provided."
        ASSERTNAL(NPARA_PORT)
        conStr = conStr & ":" & opt.Item(NPARA_PORT)
    end if
    if opt.Exists(NPARA_PREFIX) then
        TRC "Option - " & NPARA_PREFIX & ",  was provided."
        ASSERTNAL(NPARA_PREFIX)
        dim tmpStr2
        tmpStr2 = opt.Item(NPARA_PREFIX)
        if InStr(tmpStr2, "/") <> 1 then
            conStr = conStr & "/" 
        end if
        conStr = conStr & tmpStr2
    end if

    TRCS conStr, "conStr"
    GetConnectionString = conStr
end function

'''''''''''''''''''''    
' Removes ?params part from URI

private function StripParams(uri)
    LTRC "StripParams"
    
    dim qmpos
    
    TRCS uri, "strip uri before"  
    ASSERTBOOL Len(uri) <> 0, L_URIZERO_Message
    
    qmpos = InStr(uri, "?")
    if qmpos <> 0 then
        StripParams = Left(uri, qmpos - 1)
    else
        StripParams = uri
    end if

    TRCS StripParams, "strip uri after"      
end function

'''''''''''''''''''''    
' Enumerate helper

private function Enumerate(session, options, resUri)
    dim filter
    dim dialect
    dim e
    dim res
    dim formattedText
    dim formatOption
    
    if options.Exists(NPARA_FILTER) then
        filter = options.Item(NPARA_FILTER)
    end if
    
    if options.Exists(NPARA_DIALECT) then
        dialect = options.Item(NPARA_DIALECT)
    end if
    
    on error resume next
    set e = session.Enumerate(resUri, filter, dialect)
    if Err.Number = T_O then
        set e = session.Enumerate(resUri, filter, dialect)
        if Err.Number = T_O then
            set e = session.Enumerate(resUri, filter, dialect)
        end if
    end if
    if Err.Number <> 0 then
        stdErr.WriteLine session.Error
    end if
    ASSERTERR L_OP_Message & "enumerate" & "."
    on error goto 0
    
    if options.Exists(NPARA_TIMEOUT) then
        e.Timeout = options.Item(NPARA_TIMEOUT)
    end if

    formatOption = ""
    if(options.Exists(NPARA_FORMAT)) then
        formatOption = options.Item(NPARA_FORMAT)
    else
        formatOption = "#text"
    end if
    
    if(LCase(formatOption) <> "#text") then
        wscript.echo "<wsman:Results xmlns:wsman=""http://schemas.microsoft.com/ws/2005/03/wsman/results"">"
    end if
    
    TRC "do while"
    do while Not e.AtEndOfStream
        TRC "  Loop..."            
        
        on error resume next
        res = e.ReadItem()
        if Err.Number = T_O then
            res = e.ReadItem()
            if Err.Number = T_O then
                res = e.ReadItem()
            end if
        end if
        if Err.Number <> 0 then
            stdErr.WriteLine e.Error
            wscript.echo "</wsman:Results>"
        end if
        ASSERTERR "e.ReadItem()"
        on error goto 0
        
        'reformat if requested
        on error resume next
        err.clear
        if not Reformat(res,formattedText,formatOption) then
            exit do
        end if
        wscript.echo formattedText
    loop        
    TRC "end loop"
    
    if(LCase(formatOption) <> "#text") then
        wscript.echo "</wsman:Results>"
    end if
           
    set e = Nothing 
    Enumerate = ""
end function

private function GetSuffix(resUri)
    LTRC "GetSuffix"

    TRCS resUri, "resUri"  
    ASSERTBOOL Len(resUri) <> 0, L_URIZERO_Message

    GetSuffix = "_INPUT"
end function

'''''''''''''''''''''    
' Returns XML ns depending on the type of URI

private function GetXmlNs(resUri)
    LTRC "GetXmlNs"
    
    dim resUriLCase
    dim s1
    dim s2
    
    
    TRCS resUri, "resUri"  
    ASSERTBOOL Len(resUri) <> 0, L_URIZERO_Message

    resUriLCase = LCase(resUri)

    if InStr(resUriLCase, URI_IPMI) <> 0 then
        LTRC "IPMI"
        GetXmlNs = StripParams(resUri)
    elseif InStr(resUriLCase, URI_WMIV1) <> 0 then
        LTRC "V1"
        GetXmlNs = StripParams(resUri)
    else 
        LTRC "V2"
        GetXmlNs = StripParams(resUri) & ".xsd"
    end if
    
    GetXmlNs = "xmlns:p=""" & GetXmlNs & """"
    TRCS GetXmlNs, "xml ns"    
end function

'''''''''''''''''''''    
' Verify if target is in IPv6 format

private function IsIPv6(target, regexp)
    LTRC "IsIPv6"

    regexp.Pattern = PTRN_IPV6_S
    regexp.IgnoreCase = TRUE
    dim matches
    set matches = regexp.Execute(target)
    TRCS matches.Count, "matches.Count"    
    if matches.Count <> 0 then
        IsIPv6 = true
    else 
        IsIPv6 = false
    end if
end function

'''''''''''''''''''''    
' Extracts XML root node nm. from URI

private function GetRootNodeName(opr, resUri, actUri, regexp)
    LTRC "GetRootNodeName"

    dim uriTmp
    dim sfx
    dim s
    
    TRCS opr, "opr"    
    TRCS resUri, "resUri"    
    TRCS actUri, "actUri"        
            
    ASSERTBOOL Len(opr) <> 0, "'opr' parameter is 0 length or null"    
    
    sfx = ""
    select case opr
        case OP_PUT 
            uriTmp = resUri
        case OP_CRE
            uriTmp = resUri
        case OP_INV
            uriTmp = actUri
            sfx = GetSuffix(resUri)
        case else 
            GetRootNodeName = ""
            exit function
    end select
    ASSERTBOOL Len(uriTmp) <> 0, L_URIZERO_Message

    uriTmp = StripParams(uriTmp)

    regexp.Pattern = PTRN_URI_LAST
    regexp.IgnoreCase = TRUE
    dim matches
    set matches = regexp.Execute(uriTmp)
    ASSERTBOOL matches.Count = 1, L_NOLASTTOK_Message

    uriTmp = matches(0)
    ASSERTBOOL Len(uriTmp) <> 0, L_URIZEROTOK_Message
             
    GetRootNodeName = uriTmp & sfx
    TRCS GetRootNodeName, "root node name"
end function

'''''''''''''''''''''    
' Reads input XML from file or builds XML from @{...}

private function ProcessInput(wsman, regexp, operation, root, xmlns, options, resourceUriStr,sessionObj,inputStr)
    on error resume next    
    LTRC "ProcessInput"
    
    TRCS root, "root"
    TRCS xmlns, "xmlns"    
    
    dim cmdl
    dim matches
    dim hash
    dim m
    
    dim putStr
    dim elementList
    dim currentElement
    dim paramXmlFile
    dim tmpNode
        
    select case operation
        case OP_PUT 
        case OP_CRE
        case OP_INV
        case else 
            inputStr = ""
            ProcessInput = true
            exit function
    end select

    cmdl = wsman.CommandLine
    TRCS cmdl, "cmdl"
    TRCSTR cmdl

    regexp.Pattern = PTRN_HASH
    regexp.IgnoreCase = TRUE
    set matches = regexp.Execute(cmdl)
    TRCS matches.Count, "hashes count" 
    ASSERTBOOL matches.Count <= 1, L_ONEHASHONLY_Message & matches.Count
    
    'if there is no hash table, take from standard input
    if matches.Count = 0 then
        regexp.Pattern = PTRN_HASH_PART
        set matches = regexp.Execute(cmdl)
        ASSERTBOOL matches.Count <> 1, L_INVHASH1_Message
        TRC "no hash provided"     
        if options.Exists(NPARA_FILE) then
            inputStr = ReadFile(options.Item(NPARA_FILE))
            ProcessInput = true
        else
            inputStr = ""
            ProcessInput = true
        end if
        TRCS inputStr, "input"    
        
    
    'if there is a hash table
    else
        putStr = ""
                
        dim found
        regexp.Pattern = PTRN_HASH_LAST
        set found = regexp.Execute(cmdl)
        ASSERTBOOL found.Count = 1, L_HASHLAST_Message

        hash = matches(0)
        TRCS hash, "hash"
        regexp.Pattern = PTRN_HASH_0
        set matches = regexp.Execute(hash)
        TRCS matches.Count, "hashes0 count"     
        ASSERTBOOL matches.Count <= 1, L_ONEHASHONLY_Message & matches.Count
        
        'if the hash table is non-empty
        if matches.Count = 0 then
            regexp.Pattern = PTRN_HASH_VAL
            set matches = regexp.Execute(hash)
            TRCS matches.Count, "hash val count"     
            ASSERTBOOL matches.Count = 1, L_INVHASH1_Message
            regexp.Pattern = PTRN_HASH_TOK
            set matches = regexp.Execute(hash)
            ASSERTBOOL matches.Count > 0, L_INVHASH2_Message
            
            'make sure there are no duplicate entries: arrays not allowed here
            dim dupCheck
            set dupCheck = CreateObject("Scripting.Dictionary")
            for each m in matches
                if(dupCheck.Exists(m.Submatches(0))) then
                    stdErr.WriteLine "Error: " & L_PUT_PARAM_NOARRAY_Message & m.Submatches(0)
                    inputStr = ""
                    ProcessInput = false
                    on error goto 0
                    dupCheck = Empty
                    Exit Function
                else
                    dupCheck.Add m.Submatches(0),1
                end if
            next
            dupCheck = Empty      
            
            if operation = OP_CRE Or operation = OP_INV then
                dim nilns
                nilns = ""
                dim parameters
                parameters = ""                
                for each m in matches
                    TRCS m.Submatches.Count, "m.Submatches.Count"        
                    ASSERTBOOL m.Submatches.Count = 3, L_INVHASH3_Message
                    parameters = parameters & "<p:" & m.Submatches(0)
                    if m.Submatches(1) = "$null" then
                        parameters = parameters & " " & ATTR_NIL
                        nilns = " " & NS_XSI
                    end if
                    parameters = parameters & ">" & Escape(m.Submatches(2)) & _
                        "</p:" & m.Submatches(0) & ">"                     
                next
                
                putStr = "<p:" & root & " " & xmlns & nilns & ">" & _
                    parameters & "</p:" & root & ">"
            
            elseif operation = OP_PUT then
                
                'get the original resource xml
                on error resume next
                putStr = sessionObj.Get(resourceUriStr)
                if Err.Number <> 0 then
                    stdErr.WriteLine sessionObj.Error
                    inputStr = ""
                    ProcessInput = false
                    on error goto 0
                    exit function
                end if
                
                'create an MSXML DomDocument object to work with the resource xml
                on error resume next
                Set paramXmlFile = CreateObject("MSXML2.DOMDocument.3.0")
                if Err.number <> 0 then
                    stdErr.WriteLine "Error: " & L_MSXML3MISSING_Message
                    inputStr = ""
                    ProcessInput = false
                    on error goto 0
                    exit function
                end if
                paramXmlFile.async = false
                
                'load the domdocument with the resource xml   
                paramXmlFile.LoadXML(putStr)
                if (paramXmlFile.parseError.errorCode <> 0) then
                    stdErr.WriteLine "Error: " & L_XMLERROR_Message & paramXmlFile.parseError.reason
                    inputStr = ""
                    ProcessInput = false
                    on error goto 0
                    exit function
                end if
                paramXmlFile.setProperty "SelectionLanguage", "XPath" 
                
                'loop through the command-line name/value pairs
                for each m in matches
                    TRCS m.Submatches.Count, "m.Submatches.Count"        
                    ASSERTBOOL m.Submatches.Count = 3, L_INVHASH3_Message
                                
                    'find the elements matching the key
                    Dim xpathString
                    xpathString = "/*/*[local-name()=""" & m.Submatches(0) & """]"
                    Set elementList = paramXmlFile.selectNodes(xpathString)
                    
                    'make sure there is 1 - error on 0 or > 1
                    if elementList.length = 0 then
                        stdErr.WriteLine "Error: " & L_PUT_PARAM_NOMATCH_Message & m.Submatches(0)
                        inputStr = ""
                        ProcessInput = false
                        on error goto 0
                        Exit Function
                    elseif elementList.length > 1 then
                        stdErr.WriteLine "Error: " & L_PUT_PARAM_MULTIMATCH_Message & m.Submatches(0)
                        inputStr = ""
                        ProcessInput = false
                        on error goto 0
                        Exit Function                  
                    else      
                        'get the node from the list
                        Set currentElement = elementList.nextNode()
                        'make sure the node does not have anything other than 1 or less text children                    
                        if currentElement.hasChildNodes() then
                            if currentElement.childNodes.length > 1 then
                                stdErr.WriteLine "Error: " & L_PUT_PARAM_NOTATTR_Message & m.Submatches(0)
                                inputStr = ""
                                ProcessInput = false
                                on error goto 0
                                Exit Function
                            else
                                dim aNode
                                Set aNode = currentElement.childNodes.nextNode()
                                if aNode.nodeType <> NODE_TEXT then
                                    stdErr.WriteLine "Error: " & L_PUT_PARAM_NOTATTR_Message & m.Submatches(0)
                                    inputStr = ""
                                    ProcessInput = false
                                    on error goto 0
                                    Exit Function
                                end if
                            end if
                        end if
                        if m.Submatches(1) = "$null" then
                            Set tmpNode = paramXmlFile.createNode(NODE_ATTRIBUTE,ATTR_NIL_NAME,NS_XSI_URI)
                            tmpNode.text = "true"
                            currentElement.setAttributeNode tmpNode
                            currentElement.text = ""
                        else
                            'delete nil attribute if present
                            currentElement.attributes.removeNamedItem(ATTR_NIL_NAME)
                            currentElement.text = Escape(m.Submatches(2))
                        end if
                    end If        
                next
                putStr = paramXmlFile.xml
            end if   
        'if the hash table is empty
        else
            if operation = OP_PUT then
                stdErr.WriteLine "Error: " & L_PUT_PARAM_EMPTY_Message
                inputStr = ""
                ProcessInput = false
                on error goto 0
                Exit Function
            else
                putStr = "<p:" & root & " " & xmlns & nilns & ">" & "</p:" & root & ">"
                ProcessInput = true
            end if
        end if

        inputStr = putStr
        ProcessInput = true
        
        TRCS inputStr, "putStr"
    end if
    on error goto 0
    exit function
end function

'''''''''''''''''''''    
' Escapes non XML chars

private function Escape(str)
    LTRC "Escape"
    
    TRC "str in [" & str & "]"

    dim i 
    for i = 1 to Len(str)
        select case Mid(str, i, 1)
            case "&" 
                Escape = Escape & "&amp;"
            case "<"
                Escape = Escape & "&lt;"
            case ">"
                Escape = Escape & "&gt;"
            case """"
                Escape = Escape & "&quot;"
            case "'"
                Escape = Escape & "&apos;"
            case else 
                Escape = Escape & Mid(str, i, 1)
        end select
    next
   
    TRC "str out [" & Escape & "]"
end function

'''''''''''''''''''''    
' Checks if this script is running under cscript.exe

private function IsCScriptEnv()
    LTRC "IsCScriptEnv"
    
    if InStrRev(LCase(WScript.FullName), "cscript.exe", -1) <> 0 then
        IsCScriptEnv = True
    else 
        IsCScriptEnv = False
    end if
end function

private function CreateSession(wsman, conStr, optDic)
    LTRC "CreateSession"

    TRCS conStr, "conStr"

    dim sessionFlags
    dim conOpt 
    dim session
    dim authVal
    dim encodingVal
    dim pw
    dim tout
    
    sessionFlags = 0
    set conOpt = Nothing

    if optDic.Exists(NPARA_ENCODING) then
        TRC "Option - " & NPARA_ENCODING & ",  was provided."
        ASSERTNAL(NPARA_ENCODING)
        ASSERTBOOL optDic.Exists(NPARA_MACHINE), "The '-encoding' option is only valid when used with the '-machine' option"
        encodingVal = optDic.Item(NPARA_ENCODING)
        TRCS encodingVal, "encodingVal"
        if LCase(encodingVal) = "utf-8" then
            sessionFlags = sessionFlags OR WSManFlagEncodingUTF8
        elseif LCase(encodingVal) <> "utf-16" then    ' default is UTF-16, so no extra flags needed
            ' Invalid!  
            ASSERTBOOL false, "The specified encoding flag is invalid."
        end if
    end if

    if not optDic.Exists(NPARA_MACHINE) then
        ASSERTBOOL not optDic.Exists(NPARA_AUTH),     "The '-" & NPARA_AUTH & "' option is only valid when used with the '-machine' option"
        ASSERTBOOL not optDic.Exists(NPARA_USERNAME), "The '-" & NPARA_USERNAME & "' option is only valid when used with the '-machine' option"
        ASSERTBOOL not optDic.Exists(NPARA_PASSWORD), "The '-" & NPARA_PASSWORD & "' option is only valid when used with the '-machine' option"
        ASSERTBOOL not optDic.Exists(NPARA_NOCACHK),  "The '-" & NPARA_NOCACHK & "' option is only valid when used with the '-machine' option"
        ASSERTBOOL not optDic.Exists(NPARA_NOCNCHK),  "The '-" & NPARA_NOCNCHK & "' option is only valid when used with the '-machine' option"
        ASSERTBOOL not optDic.Exists(NPARA_TRANSPORT), "The '-" & NPARA_TRANSPORT & "' option is only valid when used with the '-machine' option"
        ASSERTBOOL not optDic.Exists(NPARA_SPNPORT),  "The '-" & NPARA_SPNPORT & "' option is only valid when used with the '-machine' option"
    elseif optDic.Exists(NPARA_AUTH) then
        TRC "Option - " & NPARA_AUTH & ",  was provided."
        ASSERTNAL(NPARA_AUTH)
        authVal = optDic.Item(NPARA_AUTH)
        TRCS authVal, "authVal"    
        select case LCase(authVal)
            case VAL_BASIC
                'Use -username and -password.  
                ASSERTBOOL optDic.Exists(NPARA_USERNAME), "The '-" & NPARA_USERNAME & "' option must be specified for '-auth:basic'"
                sessionFlags = sessionFlags OR WSManFlagCredUsernamePassword OR WSManFlagUseBasic
            case VAL_DIGEST
                'Use -username and -password. 
                ASSERTBOOL optDic.Exists(NPARA_USERNAME), "The '-" & NPARA_USERNAME & "' option must be specified for '-auth:digest'"
                sessionFlags = sessionFlags OR WSManFlagCredUsernamePassword OR WSManFlagUseDigest
            case VAL_NEGOTIATE
                '-username and -password are optional. 
                sessionFlags = sessionFlags OR WSManFlagUseNegotiate
            case else 
                ASSERTBOOL false, "The specified '-" & NPARA_AUTH & "' flag '" & authVal & "' has an invalid value."
        end select
    end if
    
    if optDic.Exists(NPARA_USERNAME) then
        set conOpt = wsman.CreateConnectionOptions
        conOpt.UserName = optDic.Item(NPARA_USERNAME)
        TRCS conOpt.UserName, "conOpt.UserName"    
        if optDic.Exists(NPARA_PASSWORD) then
            conOpt.Password = optDic.Item(NPARA_PASSWORD)
        else
            set pw = CreateObject("ScriptPW.Password")
            stdErr.Write L_PSWD_Message
            conOpt.Password = pw.GetPassword()
        end if
        sessionFlags = sessionFlags OR WSManFlagCredUsernamePassword
    end if

    if optDic.Exists(NPARA_NOCACHK) then
        TRC "Option - " & NPARA_NOCACHK & ",  was provided."
        'this is only valid if transport is not specified or if transport=https
        if optDic.Exists(NPARA_TRANSPORT) then
            ASSERTBOOL LCase(optDic.Item(NPARA_TRANSPORT)) = "https", "The " & NPARA_NOCACHK & "option is only valid when the transport is HTTPS"
        end if
        sessionFlags = sessionFlags OR WSManFlagSkipCACheck
    end if

    if optDic.Exists(NPARA_NOCNCHK) then
        TRC "Option - " & NPARA_NOCNCHK & ",  was provided."
        'this is only valid if transport is not specified or if transport=https
        if optDic.Exists(NPARA_TRANSPORT) then
            ASSERTBOOL LCase(optDic.Item(NPARA_TRANSPORT)) = "https", "The " & NPARA_NOCNCHK & "option is only valid when the transport is HTTPS"
        end if
        sessionFlags = sessionFlags OR WSManFlagSkipCNCheck
    end if

    if optDic.Exists(NPARA_SPNPORT) then
        TRC "Option - " & NPARA_SPNPORT & ",  was provided."
        'this is only valid if -auth is not specified or if -auth:Negotiate is specified 
        if optDic.Exists(NPARA_AUTH) then
            ASSERTBOOL LCase(optDic.Item(NPARA_AUTH)) = VAL_NEGOTIATE, "The " & NPARA_SPNPORT & " option is only valid when the authorization mechanism is " & VAL_NEGOTIATE
        end if
        sessionFlags = sessionFlags OR WSManFlagEnableSPNServerPort
    end if

    TRCS sessionFlags, "sessionFlags"    
    on error resume next
    set session = wsman.CreateSession(conStr, sessionFlags, conOpt)
    if Err.Number <> 0 then
        stdErr.WriteLine wsman.Error
    end if
    ASSERTERR "CreateSession."
    on error goto 0

    if optDic.Exists(NPARA_TIMEOUT) then
        TRC "Option - " & NPARA_TIMEOUT & ",  was provided."
        ASSERTNAL(NPARA_TIMEOUT)
        tout = optDic.Item(NPARA_TIMEOUT)
        TRCS tout, "timeout"    
        ASSERTBOOL IsNumeric(tout), "Numeric value for -timeout option is expected"
        session.Timeout = optDic.Item(NPARA_TIMEOUT)        
    end if
    
    set CreateSession = session
end function

private sub ASSERTERR(msg)
    if Err.Number <> 0 then
        if trace then
            stdErr.WriteLine "Error: [" & msg & "]"
            stdErr.WriteLine "  Number=" & Err.Number & " 0x" & hex(Err.Number)
            stdErr.WriteLine "  Description=" & Err.Description
            stdErr.WriteLine "  Source=" & Err.Source
            stdErr.WriteLine "  HelpContext=" & Err.HelpContext
            stdErr.WriteLine "  HelpFile=" & Err.HelpFile
        else
            stdErr.WriteLine "Error number: " & Err.Number & " 0x" & hex(Err.Number)
            stdErr.WriteLine Err.Description
        end if 
        WScript.Quit(ERR_GENERAL_FAILURE)
    end if
    'Err.Clear
end sub

' Assert Named Argument Length
private sub ASSERTNAL(namedArg)
    if Len(optionsDicObj.Item(namedArg)) = 0 then
        stdErr.WriteLine "Error: " & L_ARGNOVAL_Message & namedArg
        WScript.Quit(ERR_GENERAL_FAILURE)
    end if
end sub

private sub ASSERTBOOL(bool, msg)
    if Not bool then
        stdErr.WriteLine "Error: " & msg
        WScript.Quit(ERR_GENERAL_FAILURE)
    end if
end sub

private sub WRN(msg)
    if trace then
        stdErr.WriteLine "Warning: " & msg
    end if
end sub

private sub LTRC(fnc)
    if trace then
        wscript.echo ">>> " & fnc & "..."
    end if
end sub

private sub TRC(msg)
    if trace then
        wscript.echo "TRACE: " & msg
    end if
end sub

private sub TRC_(var, name)
    wscript.echo "  IsEmpty   = " & IsEmpty(var)
    wscript.echo "  IsNull    = " & IsNull(var)
    wscript.echo "  IsObject  = " & IsObject(var)
    wscript.echo "  IsArray   = " & IsArray(var)
    wscript.echo "  IsNumeric = " & IsNumeric(var)
end sub

private sub TRCOBJ(var, name)
    if trace then
        wscript.echo "TRACEOBJ: [" & name & "]"
        TRC_ var, name
    end if
end sub

private sub TRCS(var, msg)
    if trace then
        wscript.echo msg & ": [" & var & "]"
    end if
end sub

private sub TRCSTR(var)
    if trace then
        wscript.echo "string: [" & var & "] length: " & Len(var)
        dim i 
        for i = 1 to Len(var)
            wscript.echo "  <" & Mid(var, i, 1) & "> - " & AscW(Mid(var, i, 1))
        next
    end if
end sub

private sub TRCVAR(var, name)
    if trace then
        wscript.echo "TRCVAR: " & name
        wscript.echo "  Value     = [" & var & "]"
        TRC_ var, name
    end if
end sub
        
private sub Usage(stream)
dim NL 
NL = vbNewLine

stream.WriteLine _
L_HELP_000_0_Message  & NL & _
L_HELP_002_0_Message  & NL & _
L_HELP_003_0_Message  & NL & _
L_HELP_004_0_Message  & NL & _
L_HELP_005_0_Message  & NL & _
L_HELP_006_0_Message  & NL & _
L_HELP_007_0_Message  & NL & _
X_HELP_008_0_Message  & NL & _
L_HELP_009_0_Message  & NL & _
L_HELP_010_0_Message  & NL & _
L_HELP_011_0_Message  & NL & _
L_HELP_012_0_Message  & NL & _
L_HELP_015_0_Message  & NL & _
X_HELP_016_0_Message  & NL & _
L_HELP_017_0_Message  & NL & _
L_HELP_023_0_Message  & NL & _
L_HELP_024_0_Message  & NL & _
L_HELP_025_0_Message  & NL & _
L_HELP_026_0_Message  & NL & _
X_HELP_027_0_Message  & NL & _
L_HELP_028_0_Message  & NL & _
L_HELP_029_0_Message  & NL & _
L_HELP_030_0_Message  & NL & _
L_HELP_031_0_Message  & NL & _
X_HELP_032_0_Message  & NL & _
L_HELP_033_0_Message  & NL & _
L_HELP_034_0_Message  & NL & _
L_HELP_035_0_Message  & NL & _
X_HELP_036_0_Message  & NL & _
L_HELP_037_0_Message  & NL & _
L_HELP_038_0_Message  & NL & _
L_HELP_039_0_Message  & NL & _
X_HELP_040_0_Message  & NL & _
L_HELP_041_0_Message  & NL & _
L_HELP_042_0_Message  & NL & _
L_HELP_043_0_Message  & NL & _
X_HELP_044_0_Message  & NL & _
L_HELP_045_0_Message  & NL & _
L_HELP_046_0_Message  & NL & _
L_HELP_047_0_Message  & NL & _
X_HELP_048_0_Message  & NL & _
L_HELP_049_0_Message  & NL & _
L_HELP_050_0_Message  & NL & _
L_HELP_051_0_Message  & NL & _
L_HELP_052_0_Message  & NL & _
X_HELP_065_0_Message  & NL & _
L_HELP_066_0_Message  & NL & _
L_HELP_067_0_Message  & NL & _
L_HELP_068_0_Message  & NL & _
L_HELP_069_0_Message  & NL & _
L_HELP_070_0_Message  & NL & _
L_HELP_071_0_Message  & NL & _
L_HELP_072_0_Message  & NL & _
L_HELP_073_0_Message  & NL & _
L_HELP_074_0_Message  & NL & _
L_HELP_075_0_Message  & NL & _
L_HELP_076_0_Message  & NL & _
L_HELP_077_0_Message  & NL & _
L_HELP_078_0_Message  & NL & _
X_HELP_079_0_Message  & NL & _
L_HELP_080_0_Message  & NL & _
L_HELP_081_0_Message  & NL & _
L_HELP_082_0_Message  & NL & _
L_HELP_083_0_Message  & NL & _
L_HELP_084_0_Message  & NL & _
L_HELP_085_0_Message  & NL & _
X_HELP_086_0_Message  & NL & _
X_HELP_087_0_Message  & NL & _
X_HELP_088_0_Message  & NL & _
X_HELP_089_0_Message  & NL & _
L_HELP_090_0_Message  & NL & _
L_HELP_091_0_Message  & NL & _
X_HELP_092_0_Message  & NL & _
L_HELP_093_0_Message  & NL & _
X_HELP_094_0_Message  & NL & _
L_HELP_095_0_Message  & NL & _
L_HELP_096_0_Message  & NL & _
L_HELP_097_0_Message  & NL & _
L_HELP_098_0_Message  & NL & _
L_HELP_099_0_Message  & NL & _
L_HELP_100_0_Message  & NL & _
L_HELP_101_0_Message  & NL & _
L_HELP_102_0_Message  & NL & _
L_HELP_103_0_Message  & NL & _
L_HELP_104_0_Message  & NL & _
X_HELP_105_0_Message  & NL & _
L_HELP_106_0_Message  & NL & _
L_HELP_107_0_Message  & NL & _
L_HELP_108_0_Message  & NL & _
L_HELP_109_0_Message  & NL & _
L_HELP_110_0_Message  & NL & _
X_HELP_111_0_Message  & NL & _
X_HELP_112_0_Message  & NL & _
X_HELP_113_0_Message  & NL & _
L_HELP_114_0_Message  & NL & _
X_HELP_115_0_Message  & NL & _
L_HELP_116_0_Message  & NL & _
L_HELP_117_0_Message  & NL & _
L_HELP_118_0_Message  & NL & _
L_HELP_119_0_Message  & NL & _
X_HELP_120_0_Message  & NL & _
L_HELP_121_0_Message  & NL & _
X_HELP_122_0_Message  & NL & _
L_HELP_123_0_Message  & NL & _
L_HELP_124_0_Message  & NL & _
L_HELP_125_0_Message  & NL & _
L_HELP_126_0_Message  & NL & _
L_HELP_127_0_Message  & NL & _
L_HELP_128_0_Message  & NL & _
X_HELP_129_0_Message  & NL & _
L_HELP_130_0_Message  & NL & _
L_HELP_131_0_Message  & NL & _
L_HELP_132_0_Message  & NL & _
L_HELP_133_0_Message  & NL & _
X_HELP_134_0_Message  & NL & _
L_HELP_135_0_Message  & NL & _
L_HELP_136_0_Message  & NL & _
X_HELP_137_0_Message  & NL & _
L_HELP_138_0_Message  & NL & _
L_HELP_139_0_Message  & NL & _
L_HELP_140_0_Message  & NL & _
L_HELP_141_0_Message  & NL & _
L_HELP_142_0_Message  & NL & _
L_HELP_143_0_Message  & NL & _
X_HELP_144_0_Message  & NL & _
L_HELP_145_0_Message  & NL & _
L_HELP_146_0_Message  & NL & _
L_HELP_147_0_Message  & NL & _
L_HELP_148_0_Message  & NL & _
L_HELP_149_0_Message  & NL & _
L_HELP_150_0_Message  & NL & _
X_HELP_151_0_Message  & NL & _
X_HELP_152_0_Message  & NL & _
X_HELP_153_0_Message  & NL & _
L_HELP_155_0_Message  & NL & _
X_HELP_156_0_Message  & NL & _
L_HELP_157_0_Message  & NL & _
L_HELP_158_0_Message  & NL & _
L_HELP_159_0_Message  & NL & _
L_HELP_160_0_Message  & NL & _
L_HELP_161_0_Message  & NL & _
L_HELP_162_0_Message  & NL & _
X_HELP_163_0_Message  & NL & _
L_HELP_164_0_Message  & NL & _
L_HELP_165_0_Message  & NL & _
L_HELP_166_0_Message  & NL & _
L_HELP_167_0_Message  & NL & _
X_HELP_188_0_Message  & NL & _
L_HELP_189_0_Message  & NL & _
L_HELP_190_0_Message  & NL & _
L_HELP_191_0_Message  & NL & _
L_HELP_192_0_Message  & NL & _
X_HELP_193_0_Message  & NL & _
L_HELP_194_0_Message  & NL & _
L_HELP_195_0_Message  & NL & _
L_HELP_196_0_Message  & NL & _
L_HELP_197_0_Message  & NL & _
L_HELP_198_0_Message  & NL & _
L_HELP_198_1_Message  & NL & _
X_HELP_198_2_Message  & NL & _
L_HELP_198_3_Message  & NL & _
L_HELP_198_4_Message  & NL & _
L_HELP_198_5_Message  & NL & _
X_HELP_198_6_Message  & NL & _
L_HELP_198_7_Message  & NL & _
L_HELP_198_8_Message  & NL & _
L_HELP_199_0_Message  & NL & _
X_HELP_200_0_Message  & NL & _
L_HELP_201_0_Message  & NL & _
L_HELP_202_0_Message  & NL & _
L_HELP_203_0_Message  & NL & _
L_HELP_204_0_Message  & NL & _
L_HELP_205_0_Message  & NL & _
L_HELP_206_0_Message  & NL & _
X_HELP_206_1_Message  & NL & _
L_HELP_206_2_Message  & NL & _
L_HELP_206_3_Message  & NL & _
L_HELP_206_4_Message  & NL & _
L_HELP_206_5_Message  & NL & _
X_HELP_207_0_Message  & NL & _
L_HELP_208_0_Message  & NL & _
L_HELP_209_0_Message  & NL & _
L_HELP_210_0_Message  & NL & _
L_HELP_211_0_Message  & NL & _
L_HELP_212_0_Message  & NL & _
X_HELP_213_0_Message  & NL & _
X_HELP_214_0_Message  & NL & _
L_HELP_215_0_Message  & NL & _
X_HELP_216_0_Message  & NL & _
L_HELP_217_0_Message  & NL & _
L_HELP_218_0_Message  & NL & _
L_HELP_219_0_Message  & NL & _
L_HELP_220_0_Message  & NL & _
L_HELP_221_0_Message  & NL & _
X_HELP_222_0_Message  & NL & _
L_HELP_223_0_Message  & NL & _
L_HELP_224_0_Message  & NL & _
L_HELP_225_0_Message  & NL & _
L_HELP_226_0_Message  & NL & _
X_HELP_227_0_Message  & NL & _
L_HELP_228_0_Message  & NL & _
L_HELP_229_0_Message  & NL & _
L_HELP_230_0_Message  & NL & _
L_HELP_231_0_Message  & NL & _
X_HELP_232_0_Message  & NL & _
X_HELP_233_0_Message  & NL & _
X_HELP_234_0_Message  & NL & _
L_HELP_235_0_Message  & NL & _
L_HELP_236_0_Message  & NL & _
L_HELP_237_0_Message  & NL & _
L_HELP_238_0_Message  & NL & _
L_HELP_239_0_Message  & NL & _
L_HELP_240_0_Message  & NL & _
L_HELP_241_0_Message  & NL & _
L_HELP_242_0_Message  & NL & _
L_HELP_243_0_Message  & NL & _
L_HELP_244_0_Message  & NL & _
L_HELP_245_0_Message  & NL & _
L_HELP_246_0_Message  & NL & _
L_HELP_247_0_Message  & NL & _
L_HELP_248_0_Message  & NL & _
L_HELP_249_0_Message  & NL & _
L_HELP_250_0_Message  & NL & _
L_HELP_251_0_Message  & NL & _
L_HELP_252_0_Message  & NL & _
L_HELP_253_0_Message  & NL & _
L_HELP_254_0_Message  & NL & _
L_HELP_255_0_Message  & NL & _
L_HELP_256_0_Message  & NL & _
L_HELP_257_0_Message  & NL & _
L_HELP_258_0_Message  & NL & _
L_HELP_259_0_Message  & NL & _
L_HELP_260_0_Message  & NL & _
L_HELP_261_0_Message  & NL & _
L_HELP_262_0_Message  & NL & _
L_HELP_263_0_Message  & NL & _
L_HELP_264_0_Message  & NL & _
X_HELP_265_0_Message  & NL & _
L_HELP_266_0_Message  & NL & _
L_HELP_267_0_Message  & NL & _
L_HELP_268_0_Message  & NL & _
X_HELP_269_0_Message  & NL & _
L_HELP_270_0_Message  & NL & _
L_HELP_271_0_Message  & NL & _
X_HELP_272_0_Message  & NL & _
L_HELP_273_0_Message  & NL & _
L_HELP_274_0_Message  & NL & _
L_HELP_275_0_Message  & NL & _
X_HELP_276_0_Message  & NL & _
L_HELP_277_0_Message  & NL & _
X_HELP_278_0_Message  & NL & _
L_HELP_279_0_Message  & NL & _
L_HELP_280_0_Message  & NL & _
L_HELP_281_0_Message  & NL & _
X_HELP_282_0_Message  & NL & _
L_HELP_283_0_Message  & NL & _
L_HELP_284_0_Message  & NL & _
X_HELP_285_0_Message  & NL & _
L_HELP_286_0_Message  & NL & _
L_HELP_287_0_Message  & NL & _
X_HELP_288_0_Message  & NL & _
L_HELP_289_0_Message  & NL & _
L_HELP_290_0_Message  & NL & _
L_HELP_291_0_Message  & NL & _
L_HELP_292_0_Message  & NL & _
X_HELP_293_0_Message  & NL & _
X_HELP_294_0_Message  & NL & _
X_HELP_295_0_Message  & NL & _
L_HELP_296_0_Message  & NL & _
X_HELP_297_0_Message  & NL & _
L_HELP_298_0_Message  & NL & _
L_HELP_299_0_Message  & NL & _
L_HELP_300_0_Message  & NL & _
X_HELP_301_0_Message  & NL & _
L_HELP_302_0_Message  & NL & _
L_HELP_303_0_Message  & NL & _
L_HELP_304_0_Message  & NL & _
L_HELP_305_0_Message  & NL & _
X_HELP_306_0_Message  & NL & _
L_HELP_307_0_Message  & NL & _
L_HELP_308_0_Message  & NL & _
L_HELP_309_0_Message  & NL & _
X_HELP_310_0_Message  & NL & _
L_HELP_311_0_Message  & NL & _
L_HELP_312_0_Message  & NL & _
L_HELP_313_0_Message  & NL & _
L_HELP_314_0_Message  & NL & _
L_HELP_315_0_Message  & NL & _
L_HELP_316_0_Message  & NL & _
X_HELP_317_0_Message  & NL & _
L_HELP_318_0_Message  & NL & _
L_HELP_319_0_Message  & NL & _
L_HELP_320_0_Message  & NL & _
X_HELP_321_0_Message  & NL & _
L_HELP_322_0_Message  & NL & _
X_HELP_323_0_Message  & NL & _
L_HELP_324_0_Message  & NL & _
L_HELP_325_0_Message  & NL & _
L_HELP_326_0_Message  & NL & _
X_HELP_327_0_Message  & NL & _
L_HELP_328_0_Message  & NL & _
L_HELP_329_0_Message  & NL & _
L_HELP_330_0_Message  & NL & _
X_HELP_331_0_Message  & NL & _
L_HELP_332_0_Message  & NL & _
X_HELP_333_0_Message  & NL & _
L_HELP_334_0_Message  & NL & _
L_HELP_335_0_Message  & NL & _
X_HELP_336_0_Message  & NL & _
L_HELP_337_0_Message  & NL & _
L_HELP_338_0_Message  & NL & _
X_HELP_339_0_Message  & NL & _
L_HELP_340_0_Message  & NL & _
L_HELP_341_0_Message  & NL & _
L_HELP_342_0_Message  & NL & _
X_HELP_343_0_Message  & NL & _
L_HELP_344_0_Message  & NL & _
X_HELP_345_0_Message  & NL & _
X_HELP_346_0_Message  & NL & _
X_HELP_347_0_Message  & NL & _
X_HELP_348_0_Message



end sub

private function ReFormat(rawStr,formattedStr,formatOption)
    dim xslFile
    dim xmlFile
    dim xmlFileName
    dim xslFileName 
    
    on error resume next
    err.clear
    
    if LCase(formatOption) = VAL_FORMAT_XML then
        formattedStr = rawStr
    else
        set xmlFile = CreateObject("MSXML2.DOMDOCUMENT.3.0")
        if Err.number <> 0 then
            stdErr.WriteLine L_MSXML3MISSING_Message
            on error goto 0
            ReFormat = false
            exit function
        end if
 
        set xslFile = CreateObject("MSXML2.DOMDOCUMENT.3.0")
        if Err.number <> 0 then
            stdErr.WriteLine L_MSXML3MISSING_Message
            on error goto 0
            ReFormat = false
            exit function
        end if
        
        xmlFile.async = false
        xslFile.async = false
            
        xmlFile.LoadXML(rawStr)
        if (xmlFile.parseError.errorCode <> 0) then
            stdErr.WriteLine L_XMLERROR_Message & xmlFile.parseError.reason
            on error goto 0
            ReFormat = false
            exit function
        end if
        
        if LCase(formatOption) = VAL_FORMAT_TEXT then
            If Not xslFile.load(VAL_FORMAT_TEXT_XSLT) Then
                stdErr.WriteLine L_XSLERROR_Message & VAL_FORMAT_TEXT_XSLT
                if xslFile.parseError.errorCode < 0 then
                    stdErr.WriteLine xslFile.parseError.reason
                end if
                on error goto 0
                ReFormat = false
                exit function
            end if
        elseif LCase(formatOption) = VAL_FORMAT_PRETTY then    
            If Not xslFile.load(VAL_FORMAT_PRETTY_XSLT) Then
                stdErr.WriteLine L_XSLERROR_Message & VAL_FORMAT_PRETTY_XSLT
                if xslFile.parseError.errorCode < 0 then
                    stdErr.WriteLine xslFile.parseError.reason
                end if
                on error goto 0
                ReFormat = false
                exit function
            end if
        else
            stdErr.WriteLine L_FORMATLERROR_Message & formatOption
            stdErr.WriteLine 
            on error goto 0
            ReFormat = false
            exit function
        end if
        
        formattedStr = xmlFile.transformNode (xslFile)
        if Err.number <> 0 then
            stdErr.WriteLine Err.Description
            on error goto 0
            ReFormat = false
            exit function
        end if
    end if
    ReFormat = true
end function

private function ContainsAlias(resourceURIStr)
    dim posColon
    dim posSlash
    posColon = InStr(resourceURIStr,":")
    posSlash = InStr(resourceURIStr,"/")
    if (posColon = 0) Or (posSlash <= posColon) then
        ContainsAlias = true
    else
        ContainsAlias = false
    end if
end function

private function ExpandAlias(resourceURIStr)
    dim pos
    if ContainsAlias(resourceURIStr) then
        pos = InStr(resourceURIStr,"/")
        if pos > 0 then     
            select case LCase(Mid(resourceURIStr,1,pos-1))
                case VAL_ALIAS_WMI_NAME
                    ExpandAlias = VAL_ALIAS_WMI_VALUE & Mid(resourceURIStr,pos+1)         
                case VAL_ALIAS_WSMAN_NAME
                    ExpandAlias = VAL_ALIAS_WSMAN_VALUE & Mid(resourceURIStr,pos+1)         
                case VAL_ALIAS_CIMV2_9_NAME
                    ExpandAlias = VAL_ALIAS_CIMV2_9_VALUE & Mid(resourceURIStr,pos+1)         
                case VAL_ALIAS_CIMV2_NAME
                    ExpandAlias = VAL_ALIAS_CIMV2_VALUE & Mid(resourceURIStr,pos+1)         
                case else
                    ExpandAlias = resourceURIStr
            end select 
        else
            ExpandAlias = resourceURIStr
        end if   
    else
        ExpandAlias = resourceURIStr
    end if 
end function
