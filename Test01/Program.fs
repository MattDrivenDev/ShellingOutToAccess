open Microsoft.Office.Interop
open System.Runtime.InteropServices
open System.Diagnostics
open System.Linq


[<ComVisible(true)>]
module ROT = 
    
    [<DllImport("ole32.dll")>]
    extern void CreateBindCtx(System.UInt32 reserved, System.Runtime.InteropServices.ComTypes.IBindCtx& ppbc);

    let getName like =                 

        let names = System.Collections.Generic.List<string>()

        let mutable ctx = Unchecked.defaultof<System.Runtime.InteropServices.ComTypes.IBindCtx>
        let mutable table = Unchecked.defaultof<System.Runtime.InteropServices.ComTypes.IRunningObjectTable>
        let mutable mon = Unchecked.defaultof<System.Runtime.InteropServices.ComTypes.IEnumMoniker>
        let mutable lst = Array.create 1 null

        CreateBindCtx(uint32 0, &ctx)
        ctx.GetRunningObjectTable(&table)
        table.EnumRunning(&mon)
        
        while (mon.Next(1, lst, System.IntPtr.Zero)=0) do
            let mutable displayName = ""
            lst.[0].GetDisplayName(ctx, lst.[0], &displayName)    
            names.Add(displayName)
            
        let target = names.FirstOrDefault(System.Func<string, bool>(fun s -> s.ToLower().Contains(like)))

        if target = null then None else Some target


[<ComVisible(true)>]
module Program =
        
    [<DllImport("user32.dll")>]
    extern bool SetForegroundWindow(System.IntPtr hWnd);

    [<EntryPoint>]
    let main argv = 

        let path = System.Environment.CurrentDirectory + @"\test.accdb"
        let accessRuntimePath = @"c:\program files (x86)\microsoft office\office12\msaccess.exe"
        let mutable processStartInfo = new ProcessStartInfo()
        
        processStartInfo.FileName <- accessRuntimePath
        processStartInfo.Arguments <- path
        
        let p = Process.Start(processStartInfo)
        p.WaitForInputIdle(60000) |> ignore    

        let hwnd = p.MainWindowHandle

        SetForegroundWindow(hwnd) |> ignore
        
        System.Threading.Thread.Sleep(2000);

        let name = ROT.getName "test.accdb"
        
        match name with
        | None -> failwith "game over man!"
        | Some n ->

            let mutable oAccess = Marshal.BindToMoniker(n) :?> Access.Application
    
            oAccess.DoCmd.OpenForm(
                FormName="Person Form",
                View=Access.AcFormView.acNormal,
                WindowMode=Access.AcWindowMode.acWindowNormal
            )
    
            let mutable form = oAccess.Forms.Item("Person Form")
            form.Visible <- true    

            let mutable btn = form.Controls.Item("btnTestButton") :?> Access.CommandButton
            btn.add_Click(new Access.DispCommandButtonEvents_ClickEventHandler(fun _ -> 
                System.Windows.Forms.MessageBox.Show("hello access") |> ignore       
            ))

            while true do () // just keep this applicaiton open

            0 // return an integer exit code
