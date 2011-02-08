/*
iectl: JS scriptable IE control
Author: Will Timoney
Created: 11/23/2010

TODO:
	SendKeys / MouseEvent
	JSClick
*/

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using mshtml;

namespace iectl {

	public class ThreadRunner {
		public ThreadRunner(ThreadStart proc) {
			Thread t = new Thread(proc);
			t.SetApartmentState(ApartmentState.STA);
			t.Start();
			t.Join();
		}
	}

	public class Wnd {
		public const int BufSize = 0x100;

		private IntPtr handle;

		public IntPtr Handle {
			get { return handle; }
		}

		public Wnd(IntPtr _handle) {
			handle = _handle;
		}

		public static explicit operator Wnd(IntPtr handle) {
			return new Wnd(handle);
		}

		public static implicit operator IntPtr(Wnd wnd) {
			return wnd.Handle;
		}

		public delegate bool EnumWindowProc(IntPtr hWnd, IntPtr lpParam);
		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int EnumChildWindows(IntPtr hWndParent, EnumWindowProc Callback, IntPtr lpParam);
		private static bool EnumProc(IntPtr hWnd, IntPtr lpParam) {
			Wnd wnd = (Wnd)hWnd;
			((List<Wnd>)GCHandle.FromIntPtr(lpParam).Target).Add((Wnd)hWnd);
			return true;
		}

		public Wnd[] Children {
			get {
				GCHandle gch = new GCHandle();
				try {
					List<Wnd> lst = new List<Wnd>();
					gch = GCHandle.Alloc(lst);
					EnumChildWindows((IntPtr)this, new EnumWindowProc(Wnd.EnumProc), (IntPtr)gch);
					return lst.ToArray();
				}
				finally {
					if(gch.IsAllocated)
						gch.Free();
				}
			}
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
		public string ClassName {
			get {
				StringBuilder buf = new StringBuilder(BufSize);
				GetClassName((IntPtr)this, buf, buf.Capacity);
				return buf.ToString();
			}
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpWindowText, int nMaxCount);
		public string WindowText {
			get {
				StringBuilder buf = new StringBuilder(BufSize);
				GetWindowText((IntPtr)this, buf, buf.Capacity);
				return buf.ToString();
			}
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern bool IsWindowVisible(IntPtr hWnd);
		public bool IsVisible {
			get {
				return IsWindowVisible((IntPtr)this);
			}
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int GetWindowThreadProcessId(IntPtr hWnd, ref int intProcessId);
		public int ProcessId {
			get {
				int id = 0;
				GetWindowThreadProcessId((IntPtr)this, ref id);
				return id;
			}
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int SetForegroundWindow(IntPtr hWnd);
		public void SetForeground() {
			SetForegroundWindow((IntPtr)this);
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int ShowWindow(IntPtr hWnd, int nCmdShow);
		public const int SW_MAXIMIZE = 3;
		public const int SW_MINIMIZE = 6;
		public const int SW_RESTORE = 9;
		public void ShowWindow(int nCmdShow) {
			ShowWindow((IntPtr)this, nCmdShow);
		}

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern IntPtr FindWindowEx(IntPtr hWndParent, IntPtr hWndChildAfter, string lpszClass, string lpszWindow);

		[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int SendMessage(IntPtr hWnd, int iMsg, int wParam, int lParam);
		public const int WM_LBUTTONDOWN = 0x201;
		public const int WM_LBUTTONUP = 0x202;
		public const int WM_CLOSE = 0x10;
	}

	public class IEDoc {
		public const int SMTO_ABORTIFHUNG = 0x02;

		private HTMLDocument doc;

		public HTMLDocument Doc {
			get { return doc; }
		}

		public IEDoc(HTMLDocument _doc) {
			doc = _doc;
		}

		public object RunJS(string js) {
			const string proxyName = "__iectl_js_proxy__";

			object ret = null;

			new ThreadRunner(delegate() {
				Doc.parentWindow.execScript(string.Format("function {0}(){{ {1} }};", proxyName, js), "JScript");

				ret = Doc.Script.GetType().InvokeMember(proxyName, BindingFlags.InvokeMethod, null, doc.Script, null);
			});

			return ret;
		}

		public static IEDoc FromIE(Wnd ie) {
			foreach(Wnd cur in ie.Children)
				if(cur.ClassName == "Internet Explorer_Server") {
					HTMLDocument doc = GetHTMLDocument((IntPtr)cur);

					if(doc == null || doc.readyState != "complete")
						return null;

					return new IEDoc(doc);
				}

			return null;
		}

		public static IEDoc FromTitle(string pattern, out Wnd ie) {
			Regex rx = new Regex(pattern, RegexOptions.Compiled);

			ie = null;
			foreach(Wnd cur in new Wnd(IntPtr.Zero).Children)
				if((cur.ClassName == "IEFrame") && rx.Match(cur.WindowText).Success)
					return FromIE(ie = cur);

			return null;
		}

		public static IEDoc FromTitle(string pattern) {
			Wnd ie;
			return FromTitle(pattern, out ie);
		}

		public static IEDoc FromProcessId(int id, out Wnd ie) {
			ie = null;
			foreach(Wnd cur in new Wnd(IntPtr.Zero).Children)
				if(cur.ProcessId == id && cur.ClassName == "IEFrame")
					return FromIE(ie = cur);

			return null;
		}

		[DllImport("User32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		public static extern int RegisterWindowMessageA(string lpString);

		[DllImport("User32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		public static extern int SendMessageTimeoutA(IntPtr hWnd, int Msg, int wParam, int lParam, int fuFlags, int uTimeout, ref int lpdwResult);

		[DllImport("Oleacc.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		public static extern int ObjectFromLresult(int lResult, ref Guid riid, int wParam, ref HTMLDocument ppvObject);

		public static HTMLDocument GetHTMLDocument(IntPtr ieframe) {
			int WM_HTML_GETOBJECT = RegisterWindowMessageA("WM_HTML_GETOBJECT");
			if(WM_HTML_GETOBJECT == 0)
				return null;

			int docRef = 0;
			if(SendMessageTimeoutA(ieframe, WM_HTML_GETOBJECT, 0, 0, SMTO_ABORTIFHUNG, 1000, ref docRef) == 0)
				return null;

			HTMLDocument doc = null;
			Guid GUID_IHTMLDocument = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
			ObjectFromLresult(docRef, ref GUID_IHTMLDocument, 0, ref doc);

			return doc;
		}
	}

	public class IEControllerOptions {
		public int MaxSpin = 200 * 5;
		public int SleepInterval = 300;
		public string IEPath = "iexplore.exe";
	}

	[Guid("F91246BD-A003-4015-B3C1-66FFC31A3940"), ComVisible(true)]
	public class IEController {
		private Wnd window = null;

		private IEControllerOptions options = null;

		public Wnd Window {
			get { return window; }
		}

		public IEControllerOptions Options {
			get { return options; }
		}

		public IEController(IEControllerOptions _options) {
			options = _options;
		}

		#region Handling

		public delegate bool DocAction(IEDoc doc);

		public bool HandleDocAction(DocAction action) {
			IEDoc doc = IEDoc.FromIE(Window);

			if(doc == null)
				return false;

			return action(doc);
		}

		public delegate bool Spinner();

		public void Spin(Spinner spinner, string message) {
			if(!string.IsNullOrEmpty(message))
				Console.WriteLine("Spin: " + message);

			for(int curSpin = 0; curSpin < Options.MaxSpin; curSpin++) {
				if(spinner())
					return;

				Thread.Sleep(Options.SleepInterval);
			}

			throw new ApplicationException("Failed to execute before timeout :(");
		}

		public void SpinDocAction(DocAction action, string message) {
			Spin(delegate() { return HandleDocAction(action); }, message);
		}

		public void SpinDocAction(DocAction action) {
			Spin(delegate() { return HandleDocAction(action); }, string.Empty);
		}

		public IEDoc Document {
			get {
				IEDoc ret = null;
				SpinDocAction(delegate(IEDoc doc) {
					ret = doc;
					return true;
				});
				return ret;
			}
		}

		#endregion

		#region Alerts

		private const string alertWindowClass = "#32770";
		private const string alertTextClass = "Static";

		protected Wnd GetAlert() {
			foreach(Wnd cur in new Wnd(IntPtr.Zero).Children)
				if(cur.ProcessId == window.ProcessId && cur.ClassName == alertWindowClass)
					return cur;

			return null;
		}

		public string GetAlertTxt() {
			Wnd alert = GetAlert();

			if(alert != null)
				foreach(Wnd cur in alert.Children)
					if(cur.ClassName == alertTextClass)
						return cur.WindowText;

			return null;
		}

		public bool HasAlert() {
			return GetAlert() != null;
		}

		public bool HandleAlert(string action) {
			Wnd alert = GetAlert();

			if(alert == null || string.IsNullOrEmpty(action))
				return false;

			action = action.ToLower();

			if(action == "close")
				Wnd.SendMessage(alert, Wnd.WM_CLOSE, 0, 0);
			else if(action == "cancel" || action == "ok") {
				Wnd btn = new Wnd(Wnd.FindWindowEx(alert, IntPtr.Zero, "Button", action == "cancel" ? "Cancel" : "OK"));

				if(btn == null)
					return false;

				Wnd.SendMessage(btn, Wnd.WM_LBUTTONDOWN, 0, 0);
				Thread.Sleep(250);

				Wnd.SendMessage(btn, Wnd.WM_LBUTTONUP, 0, 0);
				Thread.Sleep(250);
			}
			else
				return false;

			return true;
		}

		#endregion

		#region Spinning Helpers

		public void Launch(string url) {
			Console.WriteLine("Launching: {0}", url);

			Process p = Process.Start(Options.IEPath, url);

			/*
			IE8+ MultiProcess handling kills this :/

			Spin(delegate() {
				IEDoc.FromProcessId(p.Id, out window);

				return window != null;
			}, string.Format("Launch: {0}", url));
			*/
		}

		public bool TryFind(string pattern) {
			Wnd ieNew = null;

			if(IEDoc.FromTitle(pattern, out ieNew) == null)
				return false;

			window = ieNew;

			return true;
		}

		public void Find(string pattern) {
			Spin(delegate() {
				return TryFind(pattern);
			}, string.Format("Find: {0}", pattern));
		}

		public void Wait() {
			object doc = Document;
		}

		public IHTMLDocument2 HTMLDocument {
			get { return (IHTMLDocument2)Document.Doc; }
		}
		
		[DllImport("Oleacc.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
		public static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);
		public const int OBJID_WINDOW = 0x0;

		public IHTMLWindow2 HTMLWindow {
			get {
				return (IHTMLWindow2)HTMLDocument.parentWindow;

				/*Guid iAccessibleGuid = new Guid("618736e0-3c3d-11cf-810c-00aa00389b71");
				object ret = null;
				AccessibleObjectFromWindow(Window, OBJID_WINDOW, ref iAccessibleGuid, ref ret);
				return ret;*/
			}
		}

		public string URL {
			get {
				return HTMLDocument.url;
			}
			set {
				string dst = value;
				if(dst.StartsWith("~"))
					dst = HTMLDocument.url.Substring(0, HTMLDocument.url.LastIndexOf('/')) + dst.Substring(1);

				HTMLDocument.url = dst;
			}
		}

		public object JS(string js) {
			return Document.RunJS(js);
		}

		public void Activate() {
			Window.SetForeground();
		}

		public void Max() {
			Window.ShowWindow(Wnd.SW_MAXIMIZE);
		}

		public void Min() {
			Window.ShowWindow(Wnd.SW_MINIMIZE);
		}

		#endregion
	}

	[Guid("2474938D-D38C-4f6e-A097-6A0B38C99A2C"), ComVisible(true)]
	public class ScriptUtils {
		private MSScriptControl.ScriptControlClass script = null;

		private IEController ie = null;

		public MSScriptControl.ScriptControlClass Script {
			get { return script; }
		}

		public IEController IE {
			get { return ie; }
		}

		public ScriptUtils(MSScriptControl.ScriptControlClass _script, IEController _ie) {
			script = _script;
			ie = _ie;
		}

		public const string globalName = "__global__";

		public void initGlobal() {
			Script.Eval(string.Format("var {0} = {{}};", globalName));
		}

		public string wrapScript(string js) {
			return string.Format("{0}.main = function(){{ {1} }}; {0}.main();", globalName, js);
		}

		public object __eval(string js) {
			return Script.Eval(js);//wrapScript(js));
		}

		public string read(string path) {
			Console.WriteLine("Reading: {0}", path);

			using(StreamReader sr = new StreamReader(path))
				return sr.ReadToEnd();
		}

		public object include(string path) {
			return __eval(read(path));
		}

		public void alert(string msg) {
			MessageBox.Show(msg, "iectl");
		}

		public string input(string prompt) {
			return Microsoft.VisualBasic.Interaction.InputBox(prompt, "iectl", string.Empty, 200, 100);
		}

		public string clipboard {
			get {
				string txt = null;
				new ThreadRunner(delegate() {
					txt = Clipboard.GetText();
				});
				return txt;
			}
			set {
				new ThreadRunner(delegate() {
					Clipboard.SetText(value);
				});
			}
		}

		public void sleep(int ms) {
			Thread.Sleep(ms);
		}

		public void puts(string msg) {
			Console.WriteLine(msg);
		}

		public string gets() {
			return Console.ReadLine();
		}

		public void SendKeys(string keys) {
			System.Windows.Forms.SendKeys.Send(keys);
			System.Windows.Forms.SendKeys.Flush();
		}

		public object reval(string js) {
			return IE.JS(js);
		}

		public IHTMLDocument document {
			get {
				return IE.HTMLDocument;
			}
		}

		public IHTMLWindow2 window {
			get {
				return IE.HTMLWindow;
			}
		}

		public IOmNavigator navigator {
			get {
				return (IOmNavigator)window.navigator;
			}
		}

		public IOmHistory history {
			get {
				return (IOmHistory)window.history;
			}
		}
	}

	public class iectl {
		static void Main(string[] args) {
			try {
				MSScriptControl.ScriptControlClass script = new MSScriptControl.ScriptControlClass();
				script.Language = "JavaScript";
				script.Timeout = 600000;
				script.AllowUI = true;
				script.UseSafeSubset = false;

				IEControllerOptions opts = new IEControllerOptions();
				IEController ie = new IEController(opts);

				ScriptUtils utils = new ScriptUtils(script, ie);
				script.AddObject("utils", utils, true); //makes var c = meta[a]; call htmlwindow - W T F

				string src = null;

				if(args.Length > 1 && args[0] == "-e") {
					StringBuilder sb = new StringBuilder();

					for(int i = 1; i < args.Length; i++)
						sb.AppendLine(args[i]);

					src = sb.ToString();
				}
				else if(args.Length > 0) {
					using(StreamReader sr = new StreamReader(args[0]))
						src = sr.ReadToEnd();
				}
				else
					throw new Exception("Usage: iectl (<filename> | -e <statements...>)");
				
				script.AddObject("ie", ie, false);

				utils.initGlobal();

				utils.include("autoexec.js");

				//try
				{
					utils.__eval(src);
				}
				/*catch(Exception ex)
				{
					if(script.Error.Number > 0)
						Console.WriteLine(script.Error.);
				}*/
			}
			catch(Exception ex) {
				Console.WriteLine(ex.ToString());
			}
		}
	}
}
