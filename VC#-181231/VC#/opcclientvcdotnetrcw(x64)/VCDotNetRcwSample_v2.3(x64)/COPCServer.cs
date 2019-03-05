//==============================================================================
// TITLE: COPCServer.cs - Ver1.003
//
// CONTENTS:
//
//
// (c) Copyright 2005 Takebishi Electric Sales Corporation
// ALL RIGHTS RESERVED.
//
// DISCLAIMER:
//  This code is provided by the Takebishi Electric Sales Corporation solely
//  to assist in understanding and use of the appropriate OPC Specification(s).
//  This code is provided as-is and without warranty or support of any sort.
//
// MODIFICATION LOG:
//
// Version	Date		By		Notes
// --------	----------	------	------
// 1.0.0.1	2005/08/03	Mike	First release.
// 1.0.0.2	2006/09/20	Mike	Modified for VS2005.
// 1.0.0.3	2006/10/19	Mike	Fixed the timing of 'RemoveGroup'.
// 1.0.0.4	2009/06/01	Yoshito Kishimoto	Fixed connection setting
// 1.0.0.5	2010/06/21	Yoshito Kishimoto	Fixed for Memory Leak
// 1.0.0.6	2011/11/14	Yoshito Kishimoto	Support Shutdown event
// 1.0.0.7	2012/01/17	Yoshito Kishimoto	Support Active/InActive setting
//==============================================================================

using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Windows.Forms;
//using System.Net;
//using System.Runtime.Serialization;
using OpcRcw.Da;
using OpcRcw.Comn;
using System.Collections.Generic;

namespace VCDotNetRcwSample
{
	public enum DEF_OPCDA
	{
		VER_NONE = 0,
		VER_10,
		VER_20,
		VER_30,
	}

	public delegate void DataChangeHandler(
		int						wTransID,
		int						iItemCount,
		int[]					iClientHds,
		object[]				vValues,
		OpcRcw.Da.FILETIME[]	ftTimeStamps,
		short[]					wQualities,
		int[]					pErrors);
	public delegate void ReadCompleteHandler(
		int						wTransID,
		int						iItemCount,
		int[]					iClientHds,
		object[]				vValues,
		OpcRcw.Da.FILETIME[]	ftTimeStamps,
		short[]					wQualities,
		int[]					pErrors);
	public delegate void WriteCompleteHandler(
		int						wTransID,
		int						iItemCount,
		int[]					iClientHds,
		int[]					pErrors);
	public delegate void CancelCompleteHandler(
		int						wTransID);

	// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡	(
	// ShutDownEventHandler듫릶귩먬뙻
	public delegate void ShutDownRequestHandler(string sReadson);
	// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡	)

 //	public class COPCServer : IOPCDataCallback				    // 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡
	public class COPCServer : IOPCDataCallback, IOPCShutdown	// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡

	{
		public event DataChangeHandler		DataChange;
		public event ReadCompleteHandler	ReadComplete;
		public event WriteCompleteHandler	WriteComplete;
		public event CancelCompleteHandler	CancelComplete;
		public event ShutDownRequestHandler ShutDownRequestEvent;		// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡

		[DllImport("ole32.dll")]
		private static extern void CoCreateInstanceEx(
			ref Guid         clsid,
			[MarshalAs(UnmanagedType.IUnknown)]
			object           punkOuter,
			uint             dwClsCtx,
			[In]
			ref COSERVERINFO pServerInfo,
			uint             dwCount,
			[In, Out]
			MULTI_QI[]       pResults);

		[DllImport("oleaut32.dll")]
		private static extern void VariantClear(IntPtr pVariant);

		[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
		private struct COSERVERINFO
		{
			public uint         dwReserved1;
			[MarshalAs(UnmanagedType.LPWStr)]
			public string       pwszName;
			public IntPtr       pAuthInfo;
			public uint         dwReserved2;
		};

		[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Auto)]
		private struct MULTI_QI
		{
			public IntPtr iid;
			[MarshalAs(UnmanagedType.IUnknown)]
			public object pItf;
			public uint   hr;
		}

		private struct SERVERPARAM
		{
			public Guid		clsid;
			public string	progID;
			public string	description;
//			public string	verIndProgID;
		};

		private static readonly uint CLSCTX_INPROC_SERVER	= 0x1;
		private static readonly uint CLSCTX_INPROC_HANDLER	= 0x2;
		private static readonly uint CLSCTX_LOCAL_SERVER	= 0x4;
		private static readonly uint CLSCTX_REMOTE_SERVER	= 0x10;
		private static readonly uint CLSCTX_ALL				= CLSCTX_INPROC_SERVER | CLSCTX_INPROC_HANDLER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER;

		private static readonly Guid IID_IUnknown		= new Guid("00000000-0000-0000-C000-000000000046");
		private static readonly Guid CLSID_SERVERLIST	= new Guid("13486D51-4821-11D2-A494-3CB306C10000");
		private static readonly Guid CLSID_AE_10		= new Guid("58E13251-AC87-11d1-84D5-00608CB8A7E9");
		private static readonly Guid CLSID_BATCH_10		= new Guid("A8080DA0-E23E-11D2-AFA7-00C04F539421");
		private static readonly Guid CLSID_BATCH_20		= new Guid("843DE67B-B0C9-11d4-A0B7-000102A980B1");
		private static readonly Guid CLSID_DA_10		= new Guid("63D5F430-CFE4-11d1-B2C8-0060083BA1FB");
		private static readonly Guid CLSID_DA_20		= new Guid("63D5F432-CFE4-11d1-B2C8-0060083BA1FB");
		private static readonly Guid CLSID_DA_30		= new Guid("CC603642-66D7-48f1-B69A-B625E73652D7");
		private static readonly Guid CLSID_DX_10		= new Guid("A0C85BB8-4161-4fd6-8655-BB584601C9E0");
		private static readonly Guid CLSID_HDA_10		= new Guid("7DE5B060-E089-11d2-A5E6-000086339399");
		private static readonly Guid CLSID_XMLDA_10		= new Guid("3098EDA4-A006-48b2-A27F-247453959408");

		private IOPCServer			m_OPCServer;
		private IOPCGroupStateMgt	m_OPCGroup;
		private IOPCGroupStateMgt2	m_OPCGroup2;
		private IOPCItemMgt			m_OPCItem;
		private IConnectionPointContainer m_OPCConnPointCntnr;
		private IConnectionPoint	m_OPCConnPoint;

		private int					m_iServerGroup;
//		private int					m_iUpdateRate;
//		private int					m_iTimeBias;
//		private int					m_iDeadBand;
		private int					m_iCallBackConnection;

		private bool				m_bConnect;
		private bool				m_bAdvise;
		private DEF_OPCDA			m_OpcdaVer;

		// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡	(
		private IConnectionPoint m_OpcShutdownConnectionPoint;
		private int m_iShutdownConnectionCookie;
		// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡	)

		public COPCServer()
		{
			m_bConnect = false;
			m_bAdvise = false;
		}

		/*------------------------------------------------------
		   Connect OPC Server
		
		   (ret)   True    OK
				   False   NG
		------------------------------------------------------*/
		public bool Connect(DEF_OPCDA OpcdaVer, string sNodeName, string sSvrName, string sGrpName, int iUpdateRate)
		{
			if (m_OPCServer != null)
				return true;

			m_OpcdaVer = OpcdaVer;
			try 
			{
				// instantiate the serverlist using CoCreateInstanceEx.
                IOPCServerList2 svrList = (IOPCServerList2)CreateInstance(CLSID_SERVERLIST, sNodeName);
				Guid clsidList;
				svrList.CLSIDFromProgID(sSvrName, out clsidList);//Error
				m_OPCServer = (IOPCServer)CreateInstance(clsidList, sNodeName);

				if (m_OPCServer != null) 
				{

					IConnectionPointContainer OPCConnPointCntnr = (IConnectionPointContainer)m_OPCServer;
					Guid guidShutdown = Marshal.GenerateGuidForType(typeof(IOPCShutdown));
					OPCConnPointCntnr.FindConnectionPoint(ref guidShutdown, out m_OpcShutdownConnectionPoint);

					m_OpcShutdownConnectionPoint.Advise(this, out m_iShutdownConnectionCookie);

					if (AddGroup(sGrpName, iUpdateRate)) 
					{
						IOPCCommon m_com = (IOPCCommon)m_OPCServer;
						m_com.SetClientName("TestClient");

						m_bConnect = true;

						Marshal.ReleaseComObject(svrList);		// 1.0.0.5 10/06/21 Kishimoto	Release Com Object
						svrList = null;							// 1.0.0.5 10/06/21 Kishimoto	Release Com Object

						return true;
					}
				}

				Marshal.ReleaseComObject(svrList);		// 1.0.0.5 10/06/21 Kishimoto	Release Com Object
				svrList = null;							// 1.0.0.5 10/06/21 Kishimoto	Release Com Object

				MessageBox.Show("Cannot connect OPC Server.", "Connect");
				return false;
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "Connect");
				return false;
			}
		}

		/*------------------------------------------------------
		   Browse OPC Server
		
		   (ret)   True    OK
				   False   NG
		------------------------------------------------------*/
		private SERVERPARAM[] BrowseServer()
		{
			try 
			{
				// instantiate the serverlist using CoCreateInstanceEx.
				IOPCServerList svrList = (IOPCServerList)CreateInstance(CLSID_SERVERLIST, null);

				// convert the interface version to a guid.
				Guid catid;
				switch (m_OpcdaVer) 
				{
					case DEF_OPCDA.VER_30:
						catid = CLSID_DA_30;
						break;
					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						catid = CLSID_DA_20;
						break;
				}

				// read clsids.
				SERVERPARAM[] svrPrms = GetServerParam(svrList, catid);

//				// release enumerator object.
//				OpcCom.Interop.ReleaseServer(enumerator);
//				enumerator = null;

				return svrPrms;
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "BrowseServer");
				return null;
			}
		}

		private SERVERPARAM[] GetServerParam(IOPCServerList svrList, Guid catid)
		{
			// get list of servers in the specified specification.
            IEnumGUID enumeratorObject = null; //enumeratorObject has been changed from object to IEnumGUID
			svrList.EnumClassesOfCategories(
				1,
				new Guid[] { catid },
				0,
				null,
				out enumeratorObject);

			ArrayList prms = new ArrayList();

			IEnumGUID enumerator = enumeratorObject;

			int fetched = 0;
			Guid[] buffer = new Guid[100];
			SERVERPARAM prm;
			do
			{
				IntPtr _buffer = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(Guid)) * 100);//alloc
				try
				{
					enumerator.Next(100, buffer, out fetched);

					IntPtr pPos = _buffer;
					for (int i = 0; i < fetched; i++)
					{
                        prm.clsid = (Guid)Marshal.PtrToStructure(pPos, typeof(Guid));
						pPos = (IntPtr)(pPos.ToInt64() + Marshal.SizeOf(typeof(Guid)));
						//prm.clsid = bufferPtr[i];
						svrList.GetClassDetails(
							ref prm.clsid,
							out prm.progID,
							out prm.description);
						prms.Add(prm);
					}
				}
				catch
				{
					break;
				}
				finally
				{
                    Marshal.FreeCoTaskMem(_buffer);//alloc한 것을 풀어분다.
				}
			}
			while (fetched > 0);

			return (SERVERPARAM[])prms.ToArray(typeof(SERVERPARAM));
		}

		private static object CreateInstance(Guid clsid, string hostName)
		{
			COSERVERINFO coserverInfo = new COSERVERINFO();
			GCHandle hClsid = GCHandle.Alloc(IID_IUnknown, GCHandleType.Pinned);
			MULTI_QI[] results = new MULTI_QI[1];

			results[0].iid  = hClsid.AddrOfPinnedObject();
			results[0].pItf = null;
			results[0].hr   = 0;

			try
			{
				// check whether connecting locally or remotely.
				uint clsctx = CLSCTX_ALL;
//				uint clsctx = CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER;
//				if (hostName != null && hostName.Length > 0)
//				{
//					clsctx = CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER;
//				}

				coserverInfo.pwszName = hostName;
				// create an instance.
				CoCreateInstanceEx(ref clsid, null, clsctx, ref coserverInfo, 1, results);
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.ToString(), "CreateInstance");
				return null;
			}

			hClsid.Free();		// 1.0.0.5 10/06/21 Kishimoto	Release alloc area
			return results[0].pItf;
		}

		/*------------------------------------------------------
		   Disconnect OPC Server
		
		   (ret)   True    OK
				   False   NG
		------------------------------------------------------*/
		public bool Disconnect()
		{
			if (m_OPCServer == null)
				return true;

			int ret;

			try 
			{
				Unadvise();
				if (m_OPCGroup != null) 
				{
					ret = Marshal.ReleaseComObject(m_OPCGroup);
					m_OPCGroup = null;
				}
				if (m_OPCGroup2 != null) 
				{
					ret = Marshal.ReleaseComObject(m_OPCGroup2);
					m_OPCGroup2 = null;
				}
				if (m_OPCConnPoint != null) 
				{
					ret = Marshal.ReleaseComObject(m_OPCConnPoint);
					m_OPCConnPoint = null;
				}

				// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡	(
				if (m_iShutdownConnectionCookie != 0)
				{
					m_OpcShutdownConnectionPoint.Unadvise(m_iShutdownConnectionCookie);
				}
				m_iShutdownConnectionCookie = 0;
				if (m_OpcShutdownConnectionPoint != null)
				{
					Marshal.ReleaseComObject(m_OpcShutdownConnectionPoint);
					m_OpcShutdownConnectionPoint = null;
				}
				// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡	)

				if (m_iServerGroup != 0)	// 06/10/19 Fixed the timing of 'RemoveGroup'.
				{							//
					m_OPCServer.RemoveGroup(m_iServerGroup, 0);	//
					m_iServerGroup = 0;		//
				}							//
				ret = Marshal.ReleaseComObject(m_OPCServer);
				m_OPCServer = null;
				m_bConnect = false;
				return true;
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "Disconnect");
				m_bConnect = false;
				return false;
			}
		}

		/*------------------------------------------------------
		   Browse Server List
		
		------------------------------------------------------*/
		public void BrowseServerNameList(DEF_OPCDA opcDaVer, string hostName, ref List<string> listServerName)
		{
			// 긖?긫궻덇뿓귩뺅궥갃
			listServerName.Clear();
			int nServerCnt;
			SERVERPARAM[] svrprmList = BrowseServer(opcDaVer, hostName, out nServerCnt);

			for (int n = 0; n < nServerCnt; n++)
				listServerName.Add(svrprmList[n].progID);
			return;
		}

		/*------------------------------------------------------
		   Browse Server
		
		------------------------------------------------------*/
		private SERVERPARAM[] BrowseServer(DEF_OPCDA opcDaVer, string hostName, out int nServerCnt)
		{
			nServerCnt = 0;
			try
			{
				// initialize
				IOPCServerList svrList = (IOPCServerList)CreateInstance(CLSID_SERVERLIST, hostName);

				// convert to GUID
				Guid catid;
				switch (opcDaVer)
				{
					case DEF_OPCDA.VER_30:
						catid = CLSID_DA_30;
						break;
					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						catid = CLSID_DA_20;
						break;
				}

				// Get class id
				SERVERPARAM[] svrPrms = GetServerParam(svrList, catid, out nServerCnt);

				return svrPrms;
			}
			catch (Exception exc)
			{
				MessageBox.Show(exc.ToString(), "BrowseServer");
				return null;
			}
		}


		/*------------------------------------------------------
		   Get Server Parameter
		
		------------------------------------------------------*/
		private SERVERPARAM[] GetServerParam(IOPCServerList svrList, Guid catid, out int nServerCnt)
		{
			// Get Server List
			IEnumGUID enumeratorObject = null;
			svrList.EnumClassesOfCategories(1, new Guid[] { catid }, 0, null, out enumeratorObject);

			ArrayList prms = new ArrayList();

			IEnumGUID enumerator = enumeratorObject;

			int fetched = 0;
			Guid[] buffer = new Guid[100];
			SERVERPARAM prm;
			nServerCnt = 0;
			do
			{
				IntPtr _buffer = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(Guid)) * 100);
				try
				{
					enumerator.Next(100, buffer, out fetched);
					IntPtr pPos = _buffer;
					for (int i = 0; i < fetched; i++)
					{
						prm.clsid = (Guid)Marshal.PtrToStructure(pPos, typeof(Guid));
						pPos = (IntPtr)(pPos.ToInt64() + Marshal.SizeOf(typeof(Guid)));
						prm.progID = "";
						prm.description = "";
						svrList.GetClassDetails(ref prm.clsid, out prm.progID, out prm.description);
						prms.Add(prm);
						nServerCnt++;
					}
				}
				catch
				{
					break;
				}
				finally
				{
					Marshal.FreeCoTaskMem(_buffer);
				}
			}
			while (fetched > 0);

			return (SERVERPARAM[])prms.ToArray(typeof(SERVERPARAM));
		}

		/*------------------------------------------------------
		   Get connect status
		
		------------------------------------------------------*/
		public bool IsConnect() { return m_bConnect; }

		/*------------------------------------------------------
		   Add OPC Group
		
		   (ret)   True    OK
				   False   NG
		------------------------------------------------------*/
		public bool AddGroup(string sGrpName, int iUpdateRate)
		{
			if (m_OPCServer == null)
				return false;
			if (m_OPCGroup != null || m_iServerGroup != 0)
				return false;

			object group = null;

			bool bActive = true;
			int iClientGroup = 0;//1234;
			IntPtr ptrTimeBias = IntPtr.Zero;
			IntPtr ptrDeadBand = IntPtr.Zero;
			int iLCID = 0;

			Guid guidGroupStateMgt;
			Guid guidDataCallback;
			int iRevisedUpdateRate;
			int iKeepAliveTime = 10000;

			try 
			{
				switch (m_OpcdaVer) 
				{
					case DEF_OPCDA.VER_30:
						guidGroupStateMgt = Marshal.GenerateGuidForType(typeof(IOPCGroupStateMgt2));
						m_OPCServer.AddGroup(sGrpName,
											(bActive) ? 1 : 0,
											iUpdateRate,
											iClientGroup,
											ptrTimeBias,
											ptrDeadBand,
											iLCID,
											out m_iServerGroup,
											out iRevisedUpdateRate,
											ref guidGroupStateMgt,
											out group);
						m_OPCGroup2 = (IOPCGroupStateMgt2)group;
						m_OPCGroup2.SetKeepAlive(iKeepAliveTime, out iKeepAliveTime);

						m_OPCConnPointCntnr = (IConnectionPointContainer)m_OPCGroup2;
						guidDataCallback = Marshal.GenerateGuidForType(typeof(IOPCDataCallback));
						m_OPCConnPointCntnr.FindConnectionPoint(ref guidDataCallback, out m_OPCConnPoint);
						break;

					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						guidGroupStateMgt = Marshal.GenerateGuidForType(typeof(IOPCGroupStateMgt));
						m_OPCServer.AddGroup(sGrpName,
							(bActive) ? 1 : 0,
							iUpdateRate,
							iClientGroup,
							ptrTimeBias,
							ptrDeadBand,
							iLCID,
							out m_iServerGroup,
							out iRevisedUpdateRate,
							ref guidGroupStateMgt,
							out group);
						m_OPCGroup = (IOPCGroupStateMgt)group;

						m_OPCConnPointCntnr = (IConnectionPointContainer)m_OPCGroup;
						guidDataCallback = Marshal.GenerateGuidForType(typeof(IOPCDataCallback));
						m_OPCConnPointCntnr.FindConnectionPoint(ref guidDataCallback, out m_OPCConnPoint);
						break;
				}
				return true;
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "AddGroup");
				return false;
			}
		}

		/*------------------------------------------------------
		   Valid DataChange event
		
		------------------------------------------------------*/
		public void Advise() 
		{
			if (!m_bAdvise) 
			{
				m_OPCConnPoint.Advise(this, out m_iCallBackConnection);
			}
			m_bAdvise = true;
		}

		/*------------------------------------------------------
		   Invalid DataChange event
		
		------------------------------------------------------*/
		public void Unadvise() 
		{
			if (m_bAdvise) 
			{
				m_OPCConnPoint.Unadvise(m_iCallBackConnection);
				m_iCallBackConnection = 0;
			}
			m_bAdvise = false;
		}

		/*------------------------------------------------------
		   Get Datachange event valid status
		
		------------------------------------------------------*/
		public bool IsAdvise() { return m_bAdvise; }

		/*------------------------------------------------------
		   Add OPC Item
		
		   (ret)   True    OK
				   False   NG
		------------------------------------------------------*/
		public bool AddItem(string[] ItemName, int[] ClientHd, int[] ServerHd)
		{
			int iItemCount = ItemName.Length;
			OPCITEMDEF[] itemDef = new OPCITEMDEF[iItemCount];
			OPCITEMRESULT itemResult;
			IntPtr ppResult;
			IntPtr ppErrors;
			IntPtr posRes;
			int[] errors = new int[iItemCount];
			int i;

			for (i = 0; i < iItemCount; i++) 
			{
				itemDef[i].szItemID = ItemName[i];
				itemDef[i].bActive = 1;
				itemDef[i].hClient = ClientHd[i];
			}

			try
			{
				switch (m_OpcdaVer) 
				{
					case DEF_OPCDA.VER_30:
						m_OPCItem = (IOPCItemMgt)m_OPCGroup2;
						break;
					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						m_OPCItem = (IOPCItemMgt)m_OPCGroup;
						break;
				}
				m_OPCItem.AddItems(iItemCount, itemDef, out ppResult, out ppErrors);
				Marshal.Copy(ppErrors, errors, 0, iItemCount);
				posRes = ppResult;
				for (i = 0; i < iItemCount; i++) 
				{
					itemResult = (OPCITEMRESULT)Marshal.PtrToStructure(posRes, typeof(OPCITEMRESULT));
					if (errors[i] == 0) 
					{
						ServerHd[i] = itemResult.hServer;
					}
					Marshal.DestroyStructure(posRes, typeof(OPCITEMRESULT));		// 06/09/20 for VS2005
					posRes = (IntPtr)(posRes.ToInt32() + Marshal.SizeOf(typeof(OpcRcw.Da.OPCITEMRESULT)));  // 06/09/20 for VS2005
					//posRes = new IntPtr(posRes.ToInt32() + Marshal.SizeOf(typeof(OPCITEMRESULT)));
				}
				Marshal.FreeCoTaskMem(ppResult);
				Marshal.FreeCoTaskMem(ppErrors);
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "AddItem");
				return false;
			}
//			erase itemDef;
//			erase errors;
			return true;
		}

		/*------------------------------------------------------
		Execute SyncRead
		
		(ret)   True    OK
				False   NG
		------------------------------------------------------*/
		public bool SyncRead(OpcRcw.Da.OPCDATASOURCE DataSource, int[] ServerHd, object[] Values, OpcRcw.Da.FILETIME[] TimeStamps, short[] Qualities)
		{
			int iItemCount = ServerHd.Length;
			IOPCSyncIO OPCSyncIO;
			IOPCSyncIO2 OPCSyncIO2;
			IntPtr ppItemVal;
			IntPtr ppErrors;
			IntPtr posItem;
			int[] Errors = new int[iItemCount];
			OPCITEMSTATE ItemState;
			int i;

			try
			{
				switch (m_OpcdaVer) 
				{
					case DEF_OPCDA.VER_30:
						OPCSyncIO2 = (IOPCSyncIO2)m_OPCGroup2;
						OPCSyncIO2.Read(DataSource, iItemCount, ServerHd, out ppItemVal, out ppErrors);
						Marshal.Copy(ppErrors, Errors, 0, iItemCount);
						posItem = ppItemVal;
						for (i = 0; i < iItemCount; i++) 
						{
							ItemState = (OPCITEMSTATE)Marshal.PtrToStructure(posItem, typeof(OPCITEMSTATE));
							if (Errors[i] == 0) 
							{
								Values[i] = ItemState.vDataValue;
								TimeStamps[i] = ItemState.ftTimeStamp;
								Qualities[i] = ItemState.wQuality;
							}
							Marshal.DestroyStructure(posItem, typeof(OPCITEMSTATE));		// 05/02/08 Release memory
							//posItem = new IntPtr(posItem.ToInt32() + Marshal.SizeOf(typeof(OPCITEMSTATE)));		// 1.0.0.5 10/06/21 Kishimoto Fixed for Memory Leak
							posItem = (IntPtr)(posItem.ToInt32() + Marshal.SizeOf(typeof(OpcRcw.Da.OPCITEMSTATE)));	// 1.0.0.5 10/06/21 Kishimoto Fixed for Memory Leak
						}
						Marshal.FreeCoTaskMem(ppItemVal);
						Marshal.FreeCoTaskMem(ppErrors);
						break;
					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						OPCSyncIO = (IOPCSyncIO)m_OPCGroup;
						OPCSyncIO.Read(DataSource, iItemCount, ServerHd, out ppItemVal, out ppErrors);
						Marshal.Copy(ppErrors, Errors, 0, iItemCount);
						posItem = ppItemVal;
						for (i = 0; i < iItemCount; i++) 
						{
							ItemState = (OPCITEMSTATE)Marshal.PtrToStructure(posItem, typeof(OPCITEMSTATE));
							if (Errors[i] == 0) 
							{
								Values[i] = ItemState.vDataValue;
								TimeStamps[i] = ItemState.ftTimeStamp;
								Qualities[i] = ItemState.wQuality;
							}
							Marshal.DestroyStructure(posItem, typeof(OPCITEMSTATE));		// 05/02/08 Release memory
							//posItem = new IntPtr(posItem.ToInt32() + Marshal.SizeOf(typeof(OPCITEMSTATE)));		// 1.0.0.5 10/06/21 Kishimoto Fixed for Memory Leak
							posItem = (IntPtr)(posItem.ToInt32() + Marshal.SizeOf(typeof(OpcRcw.Da.OPCITEMSTATE)));	// 1.0.0.5 10/06/21 Kishimoto Fixed for Memory Leak
						}
						Marshal.FreeCoTaskMem(ppItemVal);
						Marshal.FreeCoTaskMem(ppErrors);
						break;
				}
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "SyncRead");
				return false;
			}
//			erase Errors
			System.GC.Collect();  //04/12/29 Revised Memory Leak 
			return true;
		}

		/*------------------------------------------------------
		Execute ReadMaxAge

		(ret)   True    OK
				False   NG
		------------------------------------------------------*/
		public bool ReadMaxAge(int MaxAge, int[] ServerHd, object[] Values, OpcRcw.Da.FILETIME[] TimeStamps, short[] Qualities)
		{
			if (m_OpcdaVer != DEF_OPCDA.VER_30) 
			{
				//MsgBox("This function is for OPCDA3.0.")
				return false;
			}

			int iItemCount = ServerHd.Length;
			IOPCSyncIO2 OPCSyncIO2;
			IntPtr ppItemVals;
			IntPtr ppQualities;
			IntPtr ppTimeStamps;
			IntPtr ppErrors;
			IntPtr posItem, posQual, posTime;
			int[] Errors = new int[iItemCount];
			int[] MaxAges = new int[iItemCount];
			OpcRcw.Da.FILETIME ftTimeStamp;
			int i;

			for (i = 0; i < iItemCount; i++) 
			{
				MaxAges[i] = MaxAge;
			}

			try
			{
				OPCSyncIO2 = (IOPCSyncIO2)m_OPCGroup2;
				OPCSyncIO2.ReadMaxAge(iItemCount, ServerHd, MaxAges, out ppItemVals, out ppQualities, out ppTimeStamps, out ppErrors);
				Marshal.Copy(ppErrors, Errors, 0, iItemCount);
				posItem = ppItemVals;
				posQual = ppQualities;
				posTime = ppTimeStamps;
				for (i = 0; i < iItemCount; i++) 
				{
					if (Errors[i] == 0) 
					{
						Values[i] = Marshal.GetObjectForNativeVariant(posItem);
						Qualities[i] = Marshal.ReadInt16(posQual);
						ftTimeStamp = (OpcRcw.Da.FILETIME)Marshal.PtrToStructure(posTime, typeof(OpcRcw.Da.FILETIME));
						TimeStamps[i] = ftTimeStamp;
					}

					VariantClear(posItem);	// 05/02/08 Release memory
					Marshal.DestroyStructure(posQual, typeof(Int16));		// 1.0.0.5 10/06/21 Kishimoto	Fixed for Memory Leak
					posItem = (IntPtr)(posItem.ToInt32() + 16);							// "16" is size of "VARIANT"
					posQual = (IntPtr)(posQual.ToInt32() + 2);							// "2" is size of "short"
					Marshal.DestroyStructure(posTime, typeof(OpcRcw.Da.FILETIME));		// 05/02/08 Release memory
					posTime = (IntPtr)(posTime.ToInt32() + Marshal.SizeOf(typeof(OpcRcw.Da.FILETIME)));		// 1.0.0.5 10/06/21 Kishimoto	Fixed for Memory Leak
				}
				Marshal.FreeCoTaskMem(ppItemVals);
				Marshal.FreeCoTaskMem(ppQualities);
				Marshal.FreeCoTaskMem(ppTimeStamps);
				Marshal.FreeCoTaskMem(ppErrors);
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "ReadMaxAge");
				return false;
			}
//			erase Errors
//			erase MaxAges
			System.GC.Collect();  //04/12/29 Revised Memory Leak 
			return true;
		}

		/*------------------------------------------------------
		Execute SyncWrite
		
		(ret)   True    OK
				False   NG
		------------------------------------------------------*/
		public bool SyncWrite(int[] ServerHd, object[] Value) 
		{
			int iItemCount = ServerHd.Length;
			IOPCSyncIO OPCSyncIO;
			IOPCSyncIO2 OPCSyncIO2;
			IntPtr ppErrors;
			int[] errors = new int[iItemCount];

			try
			{
				switch (m_OpcdaVer) 
				{
					case DEF_OPCDA.VER_30:
						OPCSyncIO2 = (IOPCSyncIO2)m_OPCGroup2;
						OPCSyncIO2.Write(iItemCount, ServerHd, Value, out ppErrors);
						Marshal.Copy(ppErrors, errors, 0, iItemCount);
						Marshal.FreeCoTaskMem(ppErrors);
						break;
					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						OPCSyncIO = (IOPCSyncIO)m_OPCGroup;
						OPCSyncIO.Write(iItemCount, ServerHd, Value, out ppErrors);
						Marshal.Copy(ppErrors, errors, 0, iItemCount);
						Marshal.FreeCoTaskMem(ppErrors);
						break;
				}
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "SyncWrite");
				return false;
			}
			//			erase errors
			System.GC.Collect();  //04/12/29 Revised Memory Leak 
			return true;
		}

		/*------------------------------------------------------
		Execute AsyncRead
		
		(ret)   True    OK
				False   NG
		------------------------------------------------------*/
		public bool AsyncRead(int wTransID, out int wCancelID, int[] ServerHd)
		{
			int iItemCount = ServerHd.Length;
			IOPCAsyncIO2 OPCAsyncIO2;
			IOPCAsyncIO3 OPCAsyncIO3;
			IntPtr ppErrors;
			int[] Errors = new int[iItemCount];

			wCancelID = 0;

			try
			{
				switch (m_OpcdaVer) 
				{
					case DEF_OPCDA.VER_30:
						OPCAsyncIO3 = (IOPCAsyncIO3)m_OPCGroup2;
						OPCAsyncIO3.Read(iItemCount, ServerHd, wTransID, out wCancelID, out ppErrors);
						//Marshal.Copy(ppErrors, Errors, 0, iItemCount);
						Marshal.FreeCoTaskMem(ppErrors);
						break;
					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						OPCAsyncIO2 = (IOPCAsyncIO2)m_OPCGroup;
						OPCAsyncIO2.Read(iItemCount, ServerHd, wTransID, out wCancelID, out ppErrors);
						//Marshal.Copy(ppErrors, Errors, 0, iItemCount);
						Marshal.FreeCoTaskMem(ppErrors);
						break;
				}
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "AsyncRead");
				return false;
			}
//			erase Errors
			System.GC.Collect();  //04/12/29 Revised Memory Leak 
			return true;
		}

		/*------------------------------------------------------
		Execute AsyncWrite
		
		(ret)   True    OK
				False   NG
		------------------------------------------------------*/
		public bool AsyncWrite(int wTransID, out int wCancelID, int[] ServerHd, object[] Value)
		{
			int iItemCount = ServerHd.Length;
			IOPCAsyncIO2 OPCAsyncIO2;
			IOPCAsyncIO3 OPCAsyncIO3;
			IntPtr ppErrors;
			int[] Errors = new int[iItemCount];

			wCancelID = 0;

			try
			{
				switch (m_OpcdaVer) 
				{
					case DEF_OPCDA.VER_30:
						OPCAsyncIO3 = (IOPCAsyncIO3)m_OPCGroup2;
						OPCAsyncIO3.Write(iItemCount, ServerHd, Value, wTransID, out wCancelID, out ppErrors);
						//Marshal.Copy(ppErrors, Errors, 0, iItemCount);
						Marshal.FreeCoTaskMem(ppErrors);
						break;
					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						OPCAsyncIO2 = (IOPCAsyncIO2)m_OPCGroup;
						OPCAsyncIO2.Write(iItemCount, ServerHd, Value, wTransID, out wCancelID, out ppErrors);
						//Marshal.Copy(ppErrors, Errors, 0, iItemCount);
						Marshal.FreeCoTaskMem(ppErrors);
						break;
				}
			}
			catch (Exception exc) 
			{ 
				MessageBox.Show(exc.ToString(), "AsyncWrite");
				return false;
			}
//			erase Errors
			System.GC.Collect();  //04/12/29 Revised Memory Leak 
			return true;
		}

		public void OnDataChange(
			int						dwTransid,
			int						hGroup,
			int						hrMasterquality,
			int						hrMastererror,
			int						dwCount,
			int[]					phClientItems,
			object[]				pvValues,
			short[]					pwQualities,
			OpcRcw.Da.FILETIME[]	pftTimeStamps,
			int[]					pErrors)
		{
			DataChange(
				dwTransid,
				dwCount,
				phClientItems,
				pvValues,
				pftTimeStamps,
				pwQualities,
				pErrors);
		}

		public void OnReadComplete(
			int						dwTransid,
			int						hGroup,
			int						hrMasterquality,
			int						hrMastererror,
			int						dwCount,
			int[]					phClientItems,
			object[]				pvValues,
			short[]					pwQualities,
			OpcRcw.Da.FILETIME[]	pftTimeStamps,
			int[]					pErrors)
		{
			ReadComplete(
				dwTransid,
				dwCount,
				phClientItems,
				pvValues,
				pftTimeStamps,
				pwQualities,
				pErrors);
		}

		public void OnWriteComplete(
			int		dwTransid,
			int		hGroup,
			int		hrMastererror,
			int		dwCount,
			int[]	phClientItems,
			int[]	pErrors)
		{
			WriteComplete(
				dwTransid,
				dwCount,
				phClientItems,
				pErrors);
		}

		public void OnCancelComplete(
			int dwTransid,
			int hGroup)
		{
			CancelComplete(
				dwTransid);
		}

		// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡	(
		//=====================================================
		//Function		: ShutdownRequest(string szReason)
		//Distribution	: 긘긿긞긣?긂깛귽긹깛긣
		//return		: 궶궢
		//=====================================================
		public void ShutdownRequest(string szReason)
		{
			ShutDownRequestEvent(szReason);
		}
		// 2011/11/14 긘긿긞긣?긂깛귽긹깛긣귩롷궚귢귡귝궎궸궥귡	)


		// 2012/01/12 긐깑?긵궻귺긏긡귻긳/뷄귺긏긡귻긳귩먛귟뫶궑	(
		//=====================================================
		//Function		: SetGroupActiveStatus(bool bSetActiveState, out string sErrMsg)
		//Distribution	: 긐깑?긵궻귺긏긡귻긳륉뫴귩먠믦
		//return		: true	먝뫏뒶뿹
		//				  false	먝뫏렪봲
		//=====================================================
		public bool SetGroupActiveStatus(bool bSetActiveState, out string sErrMsg)
		{
			sErrMsg = string.Empty;
			if (m_OPCServer == null)
			{
				sErrMsg = "No connecting";
				return false;
			}

			if ( ( (m_OPCGroup == null) && (m_OPCGroup2 == null) ) || m_iServerGroup == 0)
			{
				sErrMsg = "No connecting";
				return false;
			}

			// 귺긏긡귻긳륉뫴럚믦뾭
			IntPtr pActive = IntPtr.Zero;

			try
			{
				int nUpdateRate = 0;
				int nActive = 0;
				string sName = "";
				int nTimeBias = 0;
				float fDeadband = 0;
				int nLCID = 0;
				int nClientGroup = 0;
				int nServerGroup = 0;
				int nRevisedUpdateRate = 0;
				int nActiveState = (bSetActiveState == true) ? 1 : 0;

				// ?귽깛?렳뫬긽긾깏긄깏귺귩둴뺎궢갂nActiveState귩긓긯?(귺깛긜?긲긓?긤궳둴뺎걂
				pActive = Marshal.AllocHGlobal(sizeof(int));
				Marshal.StructureToPtr(nActiveState, pActive, false);

				switch (m_OpcdaVer)
				{
					case DEF_OPCDA.VER_30:
						{
							IOPCGroupStateMgt2 OPCGroup = (IOPCGroupStateMgt2)m_OPCGroup2;
							OPCGroup.GetState(out nUpdateRate, out nActive, out sName, out nTimeBias, out fDeadband, out nLCID, out nClientGroup, out nServerGroup);
							OPCGroup.SetState(IntPtr.Zero, out nRevisedUpdateRate, pActive, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
						}
						break;
					case DEF_OPCDA.VER_10:
					case DEF_OPCDA.VER_20:
					default:
						{
							IOPCGroupStateMgt OPCGroup = (IOPCGroupStateMgt)m_OPCGroup;
							OPCGroup.GetState(out nUpdateRate, out nActive, out sName, out nTimeBias, out fDeadband, out nLCID, out nClientGroup, out nServerGroup);
							OPCGroup.SetState(IntPtr.Zero, out nRevisedUpdateRate, pActive, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
						}
						break;
				}
				return true;
			}
			catch (Exception exc)
			{
				// 쀡둖룉뿚
				sErrMsg = exc.ToString();
				return false;
			}
			finally
			{
				// 둴뺎궢궫긽긾깏긄깏귺귩됶뺳(귺깛긜?긲긓?긤궳둴뺎궢궫궫귕걂
				Marshal.FreeHGlobal(pActive);
			}
		}
		// 2012/01/12 긐깑?긵궻귺긏긡귻긳/뷄귺긏긡귻긳귩먛귟뫶궑	)
	}
}
