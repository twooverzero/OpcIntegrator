using System.Reflection;
using System.Runtime.CompilerServices;

//
// アセンブリに関する一般情報は以下の 
// 属性セットを通して制御されます。アセンブリに関連付けられている 
// 情報を変更するには、これらの属性値を変更してください。
//
[assembly: AssemblyTitle("OPC Sample (VC#)")]
[assembly: AssemblyProduct("OPC Sample (VC#)")]
//[assembly: AssemblyDescription("This assembly is set product information.")]
[assembly: AssemblyCompany("TAKEBISHI Corporation")]
[assembly: AssemblyCopyright("Copyright(c) 2006-2012")]
//[assembly: AssemblyTrademark("DeviceXPlorer is registered trademark of Takebishi Electric")]
[assembly: AssemblyCulture("")]		
//#if Debug Then
//[assembly: AssemblyConfiguration("Debug")]
//#else
//[assembly: AssemblyConfiguration("Release")]
//#endif

//
// アセンブリのバ?ジョン情報は、以下の 4 つの属性で?成されます :
//
//      Major Version
//      Minor Version 
//      Build Number
//      Revision
//
// 下にあるように、'*' を使って、すべての値を指定するか、
// ビルドおよびリビジョン番号を既定値にすることができます。

[assembly: AssemblyVersion("2.3.0.1")]
[assembly: AssemblyFileVersion("2.3.0.1")]
//[assembly: AssemblyInformationalVersion("1.0.0.2")]

//
// アセンブリに署名するには、使用するキ?を指定しなければなりません。 
// アセンブリ署名に関する詳細については、Microsoft .NET Framework ドキュメントを参照してください。
//
// 下記の属性を使って、署名に使うキ?を制御します。 
//
// メモ : 
//   (*) キ?が指定されないと、アセンブリは署名されません。
//   (*) KeyName は、コンピュ??にインスト?ルされている
//        暗号サ?ビス プロバイ? (CSP) のキ?を?します。KeyFile は、
//       キ?を含むフ?イルです。
//   (*) KeyFile および KeyName の値が共に指定されている場合は、 
//       以下の処理が行われます :
//       (1) KeyName が CSP に見つかった場合、そのキ?が使われます。
//       (2) KeyName が存在せず、KeyFile が存在する場合、 
//           KeyFile にあるキ?が CSP にインスト?ルされ、使われます。
//   (*) KeyFile を作成するには、sn.exe (厳密な名前) ユ?ティリティを使ってください。
//       KeyFile を指定するとき、KeyFile の場所は、
//       プロジェクト出力 ディレクトリへの相対パスでなければなりません。
//       パスは、%Project Directory%\obj\<configuration> です。たとえば、KeyFile がプロジェクト ディレクトリにある場合、
//       AssemblyKeyFile 属性を 
//       [assembly: AssemblyKeyFile("..\\..\\mykey.snk")] として指定します。
//   (*) 遅延署名は高度なオプションです。
//       詳細については Microsoft .NET Framework ドキュメントを参照してください。
//
[assembly: AssemblyDelaySign(false)]
[assembly: AssemblyKeyFile("")]
[assembly: AssemblyKeyName("")]
