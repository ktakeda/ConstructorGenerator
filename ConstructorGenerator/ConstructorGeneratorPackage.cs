using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;

using EnvDTE;
using EnvDTE80;

using System.Windows.Forms;

using System.Text;
using System.Linq;
using System.Collections.Generic;
using ICSharpCode.NRefactory.CSharp;

using Microsoft.VisualStudio.TextManager.Interop;

using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Formatting;


namespace ktakeda.ConstructorGenerator
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the informations needed to show the this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidConstructorGeneratorPkgString)]
    public sealed class ConstructorGeneratorPackage : Package
    {
        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public ConstructorGeneratorPackage()
        {
            Trace.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering constructor for: {0}", this.ToString()));
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overriden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initilaization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Trace.WriteLine (string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidConstructorGeneratorCmdSet, (int)PkgCmdIDList.cmdidConstructorGenerator);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            var dte = (DTE2)GetService(typeof(SDTE));

            try
            {
                //CSharpのコードだけ処理対象
                if (dte.ActiveDocument.Language != "CSharp") return;

                //NRefactorに食べさせるために保存
                //_applicationObject.ActiveDocument.Save();

                //ActiveDocumentの全Textを取得
                var doc = dte.ActiveDocument.Object("TextDocument") as TextDocument;
                var ep = doc.StartPoint.CreateEditPoint();
                var docText = ep.GetText(doc.EndPoint);

                //NRefactorでパース
                var parser = new CSharpParser();
                SyntaxTree syntaxTree = parser.Parse(docText, "");

                //カーソル位置を取得
                TextSelection objSel = (EnvDTE.TextSelection)(dte.ActiveDocument.Selection);
                var cursor = objSel.ActivePoint;
                int line = cursor.Line;
                int column = cursor.LineCharOffset;

                //classの定義部分でなければスルー
                if (!IsAtClassDefinition(syntaxTree, line, column))
                {
                    return;
                }

                //コンストラクタを生成
                string constructorDefinition = GenerateConstructorDefinition(syntaxTree, line, column);

                //コンストラクタをinsert
                objSel.Insert(constructorDefinition, System.Convert.ToInt32(vsInsertFlags.vsInsertFlagsInsertAtStart));

                //フォーマット
                objSel.SmartFormat();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        /// <summary>
        /// (line, column)でクラス定義を行っているかどうか判定する
        /// </summary>
        /// <param name="syntaxTree"></param>
        /// <param name="line"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        private static bool IsAtClassDefinition(SyntaxTree syntaxTree, int line, int column)
        {
            var nodeAtCursor = syntaxTree.GetNodeAt(line, column);
            return nodeAtCursor is ICSharpCode.NRefactory.CSharp.TypeDeclaration;
        }

        /// <summary>
        /// クラスのコンストラクタ定義を生成する
        /// </summary>
        /// <param name="syntaxTree"></param>
        /// <param name="line"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        private static string GenerateConstructorDefinition(SyntaxTree syntaxTree, int line, int column)
        {
            var classAtCursor = syntaxTree.GetNodeAt(line, column) as ICSharpCode.NRefactory.CSharp.TypeDeclaration;

            // fieldとそのtype取得
            var dictFieldNameTypeName = new Dictionary<string, string>();
            foreach (var field in classAtCursor.Members.Where(x => x is ICSharpCode.NRefactory.CSharp.FieldDeclaration))
            {
                foreach (var variable in field.GetChildrenByRole(Roles.Variable) as AstNodeCollection<VariableInitializer>)
                {
                    dictFieldNameTypeName[variable.Name] = field.ReturnType.ToString();
                }
            }

            StringBuilder sb = new StringBuilder();

            sb.AppendFormat("public {0}(\r\n", classAtCursor.Name);
            sb.Append(String.Join(",\r\n", dictFieldNameTypeName.Select(x => x.Value + " " + x.Key)));
            sb.AppendLine(")");

            //コンストラクタのbody
            sb.AppendLine("{");
            foreach (var member in dictFieldNameTypeName)
            {
                sb.AppendFormat("this.{0} = {0};\r\n", member.Key);
            }
            sb.AppendLine("}");

            return sb.ToString();
        }
    }
}
