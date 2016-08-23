using System;
using System.Windows.Forms;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Shell;

namespace orez.opaste {

	/// <summary>
	/// Paste without moving cursor.
	/// </summary>
	internal sealed class PasteAt {

		// data
		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly Package Pkg;
		/// <summary>
		/// Service provider.
		/// </summary>
		private IServiceProvider Services {
			get { return this.Pkg; }
		}
		/// <summary>
		/// Text manager.
		/// </summary>
		private IVsTextManager Texts {
			get { return Services.GetService(typeof(SVsTextManager)) as IVsTextManager; }
		}
		/// <summary>
		/// Active text view.
		/// </summary>
		private IVsTextView TextView {
			get {
				IVsTextView v = null;
				Texts.GetActiveView(1, null, out v);
				return v;
			}
		}
		/// <summary>
		/// Active text view's buffer.
		/// </summary>
		private IVsTextLines TextBuffer {
			get {
				IVsTextLines b = null;
				TextView.GetBuffer(out b);
				return b;
			}
		}

		// constant
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		// static data
		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("c6afa260-170f-498e-8b25-424de3990614");
		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static PasteAt Instance {
			get; private set;
		}


		// constructor
		/// <summary>
		/// Initialize object, add command handlers for menu.
		/// </summary>
		/// <param name="pkg">Owner package.</param>
		private PasteAt(Package pkg) {
			if ((Pkg = pkg) == null) throw new ArgumentNullException("package");
			var srv = Services.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if (srv == null) return;
			var id = new CommandID(CommandSet, CommandId);
			var menuItem = new MenuCommand(MenuItemCallback, id);
			srv.AddCommand(menuItem);
		}


		// method
		/// <summary>
		/// Execute the command when the menu item is clicked.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private void MenuItemCallback(object sender, EventArgs e) {
			object po;
			int r = 0, c = 0;
			TextView.GetCaretPos(out r, out c);
			TextBuffer.CreateEditPoint(r, c, out po);
			EnvDTE.EditPoint p = (EnvDTE.EditPoint)po;
			p.Insert(Clipboard.GetText());
			TextView.SetCaretPos(r, c);
		}


		// static method
		/// <summary>
		/// Initialize the singleton instance of the command.
		/// </summary>
		/// <param name="pkg">Owner package.</param>
		public static void Initialize(Package pkg) {
			Instance = new PasteAt(pkg);
		}
	}
}
