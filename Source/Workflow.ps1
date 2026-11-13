# Pure workflow operations. Bundled into the standalone script by Build.ps1.
function Get-CanonicalPath {
    param([Parameter(Mandatory = $true)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    if (-not ('QBitLauncher.NativePaths' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;
namespace QBitLauncher {
    public static class NativePaths {
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern SafeFileHandle CreateFile(string path, uint access, uint share, IntPtr security, uint mode, uint flags, IntPtr template);
        [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        static extern uint GetFinalPathNameByHandle(SafeFileHandle handle, StringBuilder path, uint length, uint flags);
        public static string Resolve(string path) {
            using (SafeFileHandle handle = CreateFile(path, 0, 7, IntPtr.Zero, 3, 0x02000000, IntPtr.Zero)) {
                if (handle.IsInvalid) throw new Win32Exception(Marshal.GetLastWin32Error());
                StringBuilder value = new StringBuilder(32768);
                uint length = GetFinalPathNameByHandle(handle, value, (uint)value.Capacity, 0);
                if (length == 0 || length >= value.Capacity) throw new Win32Exception(Marshal.GetLastWin32Error());
                string result = value.ToString();
                if (result.StartsWith(@"\\?\UNC\")) return @"\\" + result.Substring(8);
                if (result.StartsWith(@"\\?\")) return result.Substring(4);
                return result;
            }
        }
    }
}
'@ -ErrorAction Stop
    }
    if ([IO.File]::Exists($full) -or [IO.Directory]::Exists($full)) {
        $full = [QBitLauncher.NativePaths]::Resolve($full)
    }
    $root = [IO.Path]::GetPathRoot($full)
    if ($full.Length -gt $root.Length) { return $full.TrimEnd([char]92, [char]47) }
    return $root
}

function Test-PathContains {
    param([string]$Parent, [string]$Child)
    if ([string]::IsNullOrWhiteSpace($Parent) -or [string]::IsNullOrWhiteSpace($Child)) { return $false }
    $parentPath = Get-CanonicalPath $Parent
    $childPath = Get-CanonicalPath $Child
    return $childPath.Equals($parentPath, [StringComparison]::OrdinalIgnoreCase) -or
        $childPath.StartsWith($parentPath.TrimEnd([char]92) + '\', [StringComparison]::OrdinalIgnoreCase)
}

function Test-CleanupRoot {
    param([string]$Path)
    $canonical = Get-CanonicalPath $Path
    if ($canonical -eq [IO.Path]::GetPathRoot($canonical)) { throw 'A drive root cannot be a cleanup target.' }
    foreach ($protected in @($env:WINDIR, $env:USERPROFILE, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $Global:ScriptDir)) {
        if ($protected -and (Test-PathContains -Parent $canonical -Child $protected)) {
            throw "This folder contains a protected application or Windows location: $Path"
        }
    }
    return $canonical
}

function Test-TreeHasReparsePoint {
    param([string]$Path)
    $pending = New-Object 'System.Collections.Generic.Stack[System.IO.DirectoryInfo]'
    $pending.Push((New-Object IO.DirectoryInfo($Path)))
    while ($pending.Count) {
        foreach ($entry in $pending.Pop().GetFileSystemInfos()) {
            if (($entry.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $true }
            if ($entry -is [IO.DirectoryInfo]) { $pending.Push($entry) }
        }
    }
    return $false
}

function Get-CleanupPlan {
    param(
        [string]$SourcePath, [string]$ExtractedPath, [string]$SelectedPath,
        [ValidateSet('SourceOnly', 'ExtractedOnly', 'Everything')][string]$Mode = 'SourceOnly'
    )
    $roots = New-Object 'System.Collections.Generic.List[string]'
    if ($Mode -in @('SourceOnly', 'Everything') -and $SourcePath -and (Test-Path -LiteralPath $SourcePath)) { $roots.Add([IO.Path]::GetFullPath($SourcePath)) }
    if ($Mode -in @('ExtractedOnly', 'Everything') -and $ExtractedPath -and (Test-Path -LiteralPath $ExtractedPath)) { $roots.Add([IO.Path]::GetFullPath($ExtractedPath)) }
    foreach ($root in $roots) { [void](Test-CleanupRoot $root) }
    $protected = @()
    if ($Mode -eq 'SourceOnly') { $protected = @($ExtractedPath, $SelectedPath) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } }
    $targets = New-Object 'System.Collections.Generic.List[object]'
    $retained = New-Object 'System.Collections.Generic.List[string]'
    $pending = New-Object 'System.Collections.Generic.Stack[string]'
    foreach ($root in $roots) { $pending.Push($root) }
    while ($pending.Count) {
        $path = $pending.Pop()
        $item = Get-Item -LiteralPath $path -Force -ErrorAction Stop
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            $retained.Add($path)
            continue
        }
        $isProtected = $false; $containsProtected = $false
        foreach ($keep in $protected) {
            if (Test-PathContains -Parent $keep -Child $path) { $isProtected = $true; break }
            if (Test-PathContains -Parent $path -Child $keep) { $containsProtected = $true }
        }
        if ($isProtected) { $retained.Add($path); continue }
        if ($containsProtected -and $item.PSIsContainer) {
            foreach ($child in (Get-ChildItem -LiteralPath $path -Force -ErrorAction Stop)) { $pending.Push($child.FullName) }
            continue
        }
        # Never recycle a directory containing a junction: enumerate down to safe
        # children so deletion cannot follow a link outside the approved tree.
        if ($item.PSIsContainer) {
            $hasLinks = Test-TreeHasReparsePoint $path
            if ($hasLinks) {
                foreach ($child in (Get-ChildItem -LiteralPath $path -Force -ErrorAction Stop)) { $pending.Push($child.FullName) }
                continue
            }
        }
        $canonical = Get-CanonicalPath $path
        if (@($targets | Where-Object { $_.CanonicalPath -eq $canonical }).Count -eq 0) {
            $targets.Add([pscustomobject]@{ Path = $path; CanonicalPath = $canonical; IsDirectory = [bool]$item.PSIsContainer })
        }
    }
    $pruned = @($targets | Where-Object {
        $candidate = $_
        @($targets | Where-Object { $_.CanonicalPath -ne $candidate.CanonicalPath -and (Test-PathContains -Parent $_.CanonicalPath -Child $candidate.CanonicalPath) }).Count -eq 0
    })
    $affectsSelected = $false
    foreach ($target in $pruned) {
        if (($SelectedPath -and (Test-PathContains $target.Path $SelectedPath)) -or
            ($ExtractedPath -and (Test-PathContains $target.Path $ExtractedPath))) { $affectsSelected = $true }
    }
    return [pscustomobject]@{ Mode = $Mode; Targets = $pruned; Retained = $retained.ToArray(); ApprovedRoots = $roots.ToArray(); RequiresAcknowledgement = $affectsSelected; SourcePath = $SourcePath; ExtractedPath = $ExtractedPath; SelectedPath = $SelectedPath }
}

function Invoke-CleanupPlan {
    param($Plan, [switch]$AcknowledgeSelectedFiles, [switch]$PermanentlyDelete)
    if ($Plan.RequiresAcknowledgement -and -not $AcknowledgeSelectedFiles) { throw 'Explicit confirmation is required to delete the extracted app or selected file.' }
    Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
    $deleted = New-Object 'System.Collections.Generic.List[string]'
    foreach ($target in $Plan.Targets) {
        if (-not (Test-Path -LiteralPath $target.Path)) { continue }
        [void](Test-CleanupRoot $target.Path)
        $current = Get-CanonicalPath $target.Path
        if ($current -ne $target.CanonicalPath) { throw "Cleanup path changed after preview: $($target.Path)" }
        if (@($Plan.ApprovedRoots | Where-Object { Test-PathContains $_ $target.Path }).Count -eq 0) { throw 'Cleanup target is outside the approved locations.' }
        $item = Get-Item -LiteralPath $target.Path -Force -ErrorAction Stop
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Cleanup will not follow a junction or symbolic link.' }
        if ($item.PSIsContainer -and (Test-TreeHasReparsePoint $target.Path)) { throw 'The folder now contains a junction or symbolic link. Prepare a new cleanup preview.' }
        if ($Plan.Mode -eq 'SourceOnly') {
            foreach ($keep in @($Plan.ExtractedPath, $Plan.SelectedPath)) {
                if ($keep -and ((Test-PathContains $target.Path $keep) -or (Test-PathContains $keep $target.Path))) { throw 'This action would remove files marked to keep.' }
            }
        }
        $recycle = [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
        if ($PermanentlyDelete) { $recycle = [Microsoft.VisualBasic.FileIO.RecycleOption]::DeletePermanently }
        if ($item.PSIsContainer) {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory($target.Path, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, $recycle, [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
        } else {
            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile($target.Path, [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, $recycle, [Microsoft.VisualBasic.FileIO.UICancelOption]::ThrowException)
        }
        $deleted.Add($target.Path)
    }
    return [pscustomobject]@{ Deleted = $deleted.ToArray(); Recycled = -not $PermanentlyDelete }
}

function ConvertTo-ShortcutLabel {
    param([string]$Label)
    $name = $Label.Trim()
    if ($name.EndsWith('.lnk', [StringComparison]::OrdinalIgnoreCase)) { $name = $name.Substring(0, $name.Length - 4) }
    if ([string]::IsNullOrWhiteSpace($name) -or $name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0 -or $name -match '[ .]$' -or
        $name -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') { throw 'Enter a valid shortcut name without reserved Windows names, trailing dots, or path characters.' }
    if ($name.Length -gt 180) { throw 'Please keep the shortcut name under 181 characters.' }
    return $name
}

function Get-AvailableShortcutPath {
    param([string]$Directory, [string]$Label)
    $name = ConvertTo-ShortcutLabel $Label
    $path = Join-Path $Directory ($name + '.lnk')
    $number = 2
    while (Test-Path -LiteralPath $path) { $path = Join-Path $Directory ("$name ($number).lnk"); $number++ }
    return $path
}

function Get-RunnableRank {
    param([IO.FileInfo]$File, [string]$RootFolder)
    $path = $File.FullName; $name = $File.BaseName
    $depth = ($path.Substring([Math]::Min($RootFolder.Length, $path.Length)).Split([char]92).Count)
    $helper = $name -match '(?i)^(?:unins\d*|uninstall.*|unitycrashhandler.*|crashpad_handler|crashreport(?:client|er).*|crashhandler.*|vcredist.*|dxsetup|cefsubprocess|redist.*|diagnostic.*|benchmark)$'
    $score = 100 - [Math]::Min(60, $depth * 5)
    $reason = 'Application'
    if ($name -match '^(?i:setup|install)(?:[._ -]|$)') { $score += 25; $reason = 'Installer' }
    if ($helper -or $path -match '(?i)[\\/](?:_CommonRedist|redist|prerequisites)[\\/]') { $score -= 150; $helper = $true; $reason = 'Helper / maintenance' }
    if ($File.Extension -in @('.bat', '.cmd')) { $score -= 5; if (-not $helper) { $reason = 'Launch script' } }
    return [pscustomobject]@{ Score = $score; Helper = $helper; Reason = $reason }
}

function New-RunnableStartInfo {
    param([string]$Path, [string]$Arguments = '', [bool]$AsAdmin = $true)
    $file = Get-Item -LiteralPath $Path -ErrorAction Stop
    if ($file.PSIsContainer -or $file.Extension -notin @('.exe', '.bat', '.cmd')) { throw 'Choose an EXE, BAT, or CMD file.' }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.UseShellExecute = $true
    $info.WorkingDirectory = $file.DirectoryName
    if ($AsAdmin) { $info.Verb = 'runas' }
    if ($file.Extension -eq '.exe') {
        $info.FileName = $file.FullName; $info.Arguments = $Arguments
    } else {
        # Elevate a Unicode PowerShell handoff, then pass the script path through
        # one environment expansion in cmd. Literal %VAR% in filenames survives.
        $payload = @{ Path = $file.FullName; Directory = $file.DirectoryName; Arguments = $Arguments } | ConvertTo-Json -Compress
        $encodedPayload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($payload))
        $command = @'
$payload = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('PAYLOAD')) | ConvertFrom-Json
$info = New-Object Diagnostics.ProcessStartInfo
$info.UseShellExecute = $false
$info.FileName = Join-Path $env:WINDIR 'System32\cmd.exe'
$info.WorkingDirectory = $payload.Directory
$info.EnvironmentVariables['QBITLAUNCHER_SCRIPT_TARGET'] = $payload.Path
$info.Arguments = '/d /v:off /s /c ""%QBITLAUNCHER_SCRIPT_TARGET%" ' + $payload.Arguments + '"'
$process = [Diagnostics.Process]::Start($info)
$process.WaitForExit()
$exitCode = $process.ExitCode
$process.Dispose()
exit $exitCode
'@
        $command = $command.Replace('PAYLOAD', $encodedPayload)
        $info.FileName = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $info.Arguments = '-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
        $info.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    }
    return $info
}

function Start-Runnable {
    param([string]$Path, [string]$Arguments = '', [bool]$AsAdmin = $true)
    $info = New-RunnableStartInfo -Path $Path -Arguments $Arguments -AsAdmin $AsAdmin
    $process = [Diagnostics.Process]::Start($info)
    if ($process) { $process.Dispose() }
}

function Initialize-FolderScanner {
    if ('QBitLauncher.FolderScan' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using System.Text;
using System.Text.RegularExpressions;
namespace QBitLauncher {
    public sealed class ArchiveOutputFile {
        public string Path;
        public long Length;
        public ArchiveOutputFile(string path,long length) { Path=path; Length=length; }
    }
    public sealed class ArchiveOutputCheck {
        public bool Valid;
        public int Checked;
        public string[] Problems;
        public string DestinationPath;
    }
    public sealed class FolderScanProgress {
        public readonly long Files;
        public readonly int Folders, PendingFolders, Runnables, Skipped;
        public readonly string CurrentPath;
        public readonly string Phase;
        public readonly bool Complete;
        public FolderScanProgress(long files, int folders, int pending, int runnables, int skipped, string path, bool complete, string phase="Scanning files") {
            Files=files; Folders=folders; PendingFolders=pending; Runnables=runnables; Skipped=skipped; CurrentPath=path; Complete=complete;
            Phase=phase;
        }
    }
    public sealed class FolderScanResult {
        public string RootFolder;
        public FileInfo[] Runnables, AllArchives, Media;
        public string[] Skipped;
        public bool Complete;
        public ArchiveOutputCheck OutputCheck;
    }
    public sealed class FolderScan : IDisposable {
        readonly string root;
        readonly string archivePath, listTool;
        readonly CancellationTokenSource cancellation = new CancellationTokenSource();
        readonly object gate = new object();
        readonly List<FileInfo> runnables = new List<FileInfo>(), archives = new List<FileInfo>(), media = new List<FileInfo>();
        readonly List<string> skipped = new List<string>();
        FolderScanProgress progress;
        int disposed;
        ArchiveOutputCheck outputCheck;
        public Task<FolderScanResult> Operation { get; private set; }
        public FolderScanProgress Progress { get { return Volatile.Read(ref progress); } }
        public FolderScan(string path) : this(path,null,null) { }
        public FolderScan(string path,string archive,string tool) {
            root=Path.GetFullPath(path);
            archivePath=archive; listTool=tool;
            progress=new FolderScanProgress(0,0,1,0,0,root,false);
            Operation=Task.Factory.StartNew<FolderScanResult>(new Func<FolderScanResult>(Scan), cancellation.Token, TaskCreationOptions.None, TaskScheduler.Default);
        }
        public FolderScanResult Snapshot(bool complete) {
            lock(gate) { return new FolderScanResult { RootFolder=root, Runnables=runnables.ToArray(), AllArchives=archives.ToArray(), Media=media.ToArray(), Skipped=skipped.ToArray(), Complete=complete, OutputCheck=outputCheck }; }
        }
        public void Cancel() { if(Volatile.Read(ref disposed)==0) cancellation.Cancel(); }
        public void Dispose() {
            if(Interlocked.Exchange(ref disposed,1)!=0) return;
            cancellation.Cancel();
            // Never wait on disk I/O from the UI thread, including cancellation.
            Operation.ContinueWith(task => { if(task.IsFaulted) { var observed=task.Exception; } cancellation.Dispose(); }, TaskScheduler.Default);
        }
        FolderScanResult Scan() {
            CancellationToken token=cancellation.Token;
            DirectoryInfo start=new DirectoryInfo(root);
            if(!start.Exists) throw new DirectoryNotFoundException("The scan folder is missing or inaccessible: "+root);
            if((start.Attributes & FileAttributes.Directory)==0) throw new IOException("Choose a folder to scan: "+root);
            Queue<DirectoryInfo> pending=new Queue<DirectoryInfo>(); pending.Enqueue(start);
            long files=0; int folders=0;
            Stopwatch pulse=Stopwatch.StartNew();
            while(pending.Count>0) {
                token.ThrowIfCancellationRequested();
                DirectoryInfo directory=pending.Dequeue();
                Publish(files,folders,pending.Count+1,directory.FullName,false);
                try {
                    // Stream metadata only. Do not read payloads or allocate an
                    // array for every file in a large directory.
                    using(IEnumerator<FileSystemInfo> entries=directory.EnumerateFileSystemInfos().GetEnumerator()) {
                        while(true) {
                            token.ThrowIfCancellationRequested();
                            if(!entries.MoveNext()) break;
                            FileSystemInfo entry=entries.Current;
                            try {
                                if((entry.Attributes & FileAttributes.ReparsePoint)!=0) { lock(gate) skipped.Add(entry.FullName); continue; }
                                DirectoryInfo child=entry as DirectoryInfo;
                                if(child!=null) { pending.Enqueue(child); continue; }
                                files++;
                                FileInfo file=(FileInfo)entry;
                                string extension=file.Extension.ToLowerInvariant();
                                if(extension==".exe" || extension==".bat" || extension==".cmd") { lock(gate) runnables.Add(file); }
                                else if(extension==".zip" || extension==".rar" || extension==".7z" || extension==".iso" || extension==".img" || file.Name.EndsWith(".7z.001",StringComparison.OrdinalIgnoreCase) || file.Name.EndsWith(".zip.001",StringComparison.OrdinalIgnoreCase)) { lock(gate) archives.Add(file); }
                                else if(extension==".mp4" || extension==".mkv" || extension==".avi" || extension==".mov" || extension==".wmv" || extension==".flv" || extension==".mp3" || extension==".flac" || extension==".wav" || extension==".aac" || extension==".ogg" || extension==".webm" || extension==".m4a") { lock(gate) media.Add(file); }
                            } catch(UnauthorizedAccessException) { lock(gate) skipped.Add(entry.FullName); }
                              catch(IOException) { lock(gate) skipped.Add(entry.FullName); }
                              catch(SecurityException) { lock(gate) skipped.Add(entry.FullName); }
                            if(pulse.ElapsedMilliseconds>=150) { Publish(files,folders,pending.Count+1,directory.FullName,false); pulse.Restart(); }
                        }
                    }
                } catch(UnauthorizedAccessException) { lock(gate) skipped.Add(directory.FullName); }
                  catch(IOException) { lock(gate) skipped.Add(directory.FullName); }
                  catch(SecurityException) { lock(gate) skipped.Add(directory.FullName); }
                folders++;
            }
            token.ThrowIfCancellationRequested();
            if(!String.IsNullOrEmpty(archivePath) && !String.IsNullOrEmpty(listTool)) {
                Publish(files,folders,0,archivePath,false,"Checking archive file list");
                ArchiveOutputFile[] manifest=ReadArchiveManifest(archivePath,listTool,15000,token);
                if(manifest!=null) {
                    outputCheck=ValidateArchiveOutput(manifest,root,token,path=>Publish(files,folders,0,path,false,"Checking extracted file sizes"));
                }
            }
            Publish(files,folders,0,root,true);
            return Snapshot(true);
        }
        void Publish(long files,int folders,int pending,string path,bool complete,string phase="Scanning files") {
            lock(gate) { Volatile.Write(ref progress,new FolderScanProgress(files,folders,pending,runnables.Count,skipped.Count,path,complete,phase)); }
        }
        public static ArchiveOutputFile[] ReadArchiveManifest(string archive,string tool,int timeout,CancellationToken token) {
            if(String.IsNullOrEmpty(tool) || !File.Exists(tool)) return null;
            using(Process process=new Process()) {
                process.StartInfo=new ProcessStartInfo { FileName=tool, Arguments="l -slt -ba -sccUTF-8 -p- -- \""+archive+"\"", UseShellExecute=false, CreateNoWindow=true, RedirectStandardOutput=true, RedirectStandardError=true, StandardOutputEncoding=Encoding.UTF8 };
                bool started=false;
                try {
                    process.Start(); started=true;
                    Task<string> output=process.StandardOutput.ReadToEndAsync(), errors=process.StandardError.ReadToEndAsync();
                    Stopwatch watch=Stopwatch.StartNew();
                    while(!process.WaitForExit(Math.Min(100,Math.Max(1,timeout)))) {
                        token.ThrowIfCancellationRequested();
                        if(watch.ElapsedMilliseconds>=timeout) return null;
                    }
                    token.ThrowIfCancellationRequested();
                    if(process.ExitCode!=0) return null;
                    while(!output.IsCompleted || !errors.IsCompleted) {
                        token.ThrowIfCancellationRequested();
                        if(watch.ElapsedMilliseconds>=timeout) return null;
                        Thread.Sleep(10);
                    }
                    string text=output.GetAwaiter().GetResult(); errors.GetAwaiter().GetResult();
                    List<ArchiveOutputFile> entries=new List<ArchiveOutputFile>();
                    foreach(string block in Regex.Split(text,@"\r?\n\s*\r?\n")) {
                        token.ThrowIfCancellationRequested();
                        Match path=Regex.Match(block,@"(?m)^Path = (.+)\r?$"), length=Regex.Match(block,@"(?m)^Size = ([0-9]+)\r?$");
                        long bytes;
                        if(!path.Success || !length.Success || !Int64.TryParse(length.Groups[1].Value,out bytes) || Regex.IsMatch(block,@"(?m)^Folder = \+\r?$|^Attributes = D")) continue;
                        entries.Add(new ArchiveOutputFile(path.Groups[1].Value.TrimEnd('\r'),bytes));
                    }
                    return entries.Count==0 ? null : entries.ToArray();
                } catch(OperationCanceledException) { throw; }
                  catch(Exception) { return null; }
                finally { if(started) { try { if(!process.HasExited) { process.Kill(); process.WaitForExit(500); } } catch(InvalidOperationException) { } } }
            }
        }
        public static ArchiveOutputCheck ValidateArchiveOutput(ArchiveOutputFile[] manifest,string destination,CancellationToken token,Action<string> onProgress) {
            string root=Path.GetFullPath(destination).TrimEnd('\\','/')+Path.DirectorySeparatorChar;
            List<string> problems=new List<string>(); int count=0;
            Stopwatch pulse=Stopwatch.StartNew();
            foreach(ArchiveOutputFile entry in manifest) {
                token.ThrowIfCancellationRequested();
                string path=Path.GetFullPath(Path.Combine(root,entry.Path.Replace('/','\\')));
                if(Path.IsPathRooted(entry.Path) || !path.StartsWith(root,StringComparison.OrdinalIgnoreCase)) throw new IOException("The archive contains a path outside the extraction destination.");
                try {
                    FileInfo file=new FileInfo(path);
                    if(!file.Exists) problems.Add("Missing: "+entry.Path);
                    else if(file.Length!=entry.Length) problems.Add(String.Format("Incomplete: {0} ({1:N0} of {2:N0} bytes)",entry.Path,file.Length,entry.Length));
                } catch(IOException) { problems.Add("Cannot check: "+entry.Path); }
                  catch(UnauthorizedAccessException) { problems.Add("Cannot check: "+entry.Path); }
                  catch(SecurityException) { problems.Add("Cannot check: "+entry.Path); }
                count++;
                if(onProgress!=null && pulse.ElapsedMilliseconds>=150) { onProgress(path); pulse.Restart(); }
            }
            return new ArchiveOutputCheck { Valid=count>0 && problems.Count==0, Checked=count, Problems=problems.ToArray(), DestinationPath=destination };
        }
    }
}
'@ -ErrorAction Stop
}

function ConvertTo-FolderInventory {
    param($ScanResult)
    $root = $ScanResult.RootFolder
    $sorted = @($ScanResult.Runnables | Sort-Object @{ Expression = { (Get-RunnableRank $_ $root).Score }; Descending = $true }, FullName)
    $payloads = @($ScanResult.AllArchives | Where-Object {
        $_.FullName -notmatch '(?i)[\\/](?:plugins|_Data|resources|assets|redist|crack)[\\/]' -and
        $_.Name -notmatch '(?i)\.part0*[2-9][0-9]*\.rar$|\.part0*1[0-9]+\.rar$'
    } | Sort-Object @{ Expression = { $_.FullName.Split([char]92).Count } }, @{ Expression = { $_.Length }; Descending = $true })
    return [pscustomobject]@{ RootFolder = $root; Runnables = $sorted; Archives = $payloads; AllArchives = @($ScanResult.AllArchives); Media = @($ScanResult.Media); Skipped = @($ScanResult.Skipped); Complete = $ScanResult.Complete; OutputCheck = $ScanResult.OutputCheck }
}

function Start-FolderScan {
    param([string]$RootFolder, [string]$ArchivePath, [string]$ListToolPath)
    Initialize-FolderScanner
    return New-Object QBitLauncher.FolderScan($RootFolder, $ArchivePath, $ListToolPath)
}

function Get-FolderInventory {
    param([string]$RootFolder)
    $scan = Start-FolderScan $RootFolder
    try { return ConvertTo-FolderInventory ($scan.Operation.GetAwaiter().GetResult()) }
    finally { $scan.Dispose() }
}

function Get-PrimaryContent {
    param($Inventory)
    $root = $Inventory.RootFolder.TrimEnd([char]92)
    $shallow = @($Inventory.Archives | Where-Object { $_.DirectoryName.TrimEnd([char]92) -eq $root })
    $mainRunnables = @($Inventory.Runnables | Where-Object { -not (Get-RunnableRank $_ $root).Helper })
    if ($mainRunnables.Count -and -not $shallow.Count) { return [pscustomobject]@{ Kind = 'Runnables'; Files = @($Inventory.Runnables) } }
    $candidates = if ($shallow.Count) { $shallow } else { @($Inventory.Archives) }
    if ($candidates.Count) {
        $depth = ($candidates[0].DirectoryName.Split([char]92).Count)
        $candidates = @($candidates | Where-Object { $_.DirectoryName.Split([char]92).Count -eq $depth })
        return [pscustomobject]@{ Kind = 'Archives'; Files = $candidates }
    }
    if ($Inventory.Runnables.Count) { return [pscustomobject]@{ Kind = 'Runnables'; Files = @($Inventory.Runnables) } }
    return [pscustomobject]@{ Kind = 'Other'; Files = @($Inventory.Media) }
}

function Get-QBittorrentRepairLines {
    param([string[]]$Lines, [int]$Port = 8080)
    if ($Port -lt 1 -or $Port -gt 65535) { throw 'API port must be between 1 and 65535.' }
    $modern = @($Lines | Where-Object { $_ -match '^\[WebUI\]\s*$' }).Count -gt 0
    $section = if ($modern) { 'WebUI' } else { 'Preferences' }
    $prefix = if ($modern) { '' } else { 'WebUI\' }
    $result = Set-QBittorrentIniValue -Lines $Lines -Section $section -Key ($prefix + 'Enabled') -Value 'true'
    $result = Set-QBittorrentIniValue -Lines $result -Section $section -Key ($prefix + 'Port') -Value ([string]$Port)
    $result = Set-QBittorrentIniValue -Lines $result -Section $section -Key ($prefix + 'Address') -Value '127.0.0.1'
    # Preserve every authentication, password, HTTPS, and host-validation setting.
    return ,$result
}

function Enable-QBittorrentLocalApiConfig {
    param([int]$Port = 8080)
    $state = Get-QBittorrentWebApiState
    if (-not $state.Exists) { throw 'qBittorrent.ini was not found. Start qBittorrent once before setting up cleanup.' }
    if (Test-QBittorrentProcessRunning) { throw 'Close qBittorrent before applying its configuration changes.' }
    $lines = [IO.File]::ReadAllLines($state.ConfigPath)
    $updated = Get-QBittorrentRepairLines -Lines $lines -Port $Port
    $backup = $state.ConfigPath + '.qBitLauncher-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.bak'
    [IO.File]::Copy($state.ConfigPath, $backup)
    $temporary = $state.ConfigPath + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [IO.File]::WriteAllLines($temporary, [string[]]$updated, (New-Object Text.UTF8Encoding($false)))
        [IO.File]::Replace($temporary, $state.ConfigPath, [NullString]::Value)
    } finally { if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) } }
    Write-LogMessage 'Enabled the qBittorrent loopback API while preserving existing authentication settings.'
    return $backup
}

function Get-RunnableIconBytes {
    param([IO.FileInfo]$File, [string]$CacheDirectory)
    try {
        $hash = [Security.Cryptography.SHA256]::Create()
        try { $key = [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($File.FullName + '|' + $File.LastWriteTimeUtc.Ticks + '|' + $File.Length))).Replace('-', '') } finally { $hash.Dispose() }
        [IO.Directory]::CreateDirectory($CacheDirectory) | Out-Null
        $cachePath = Join-Path $CacheDirectory ($key + '.png')
        if ([IO.File]::Exists($cachePath)) { return ,([IO.File]::ReadAllBytes($cachePath)) }
        $icon = [Drawing.Icon]::ExtractAssociatedIcon($File.FullName)
        if (-not $icon) { return $null }
        try {
            $bitmap = $icon.ToBitmap(); $stream = New-Object IO.MemoryStream
            try { $bitmap.Save($stream, [Drawing.Imaging.ImageFormat]::Png); $bytes = $stream.ToArray(); [IO.File]::WriteAllBytes($cachePath, $bytes); return ,$bytes }
            finally { $stream.Dispose(); $bitmap.Dispose() }
        } finally { $icon.Dispose() }
    } catch { return $null }
}

function New-WindowsShortcut {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ShortcutPath,
        [Parameter(Mandatory = $true)]
        [string]$TargetPath,
        [string]$WorkingDirectory,
        [string]$Arguments = ""
    )

    $target = Get-Item -LiteralPath $TargetPath -ErrorAction Stop
    if ($target.PSIsContainer) {
        throw "The shortcut target must be a file: $TargetPath"
    }
    if ([string]::IsNullOrWhiteSpace($WorkingDirectory)) {
        $WorkingDirectory = $target.DirectoryName
    }
    if (-not (Test-Path -LiteralPath $WorkingDirectory -PathType Container)) {
        throw "The shortcut working directory does not exist: $WorkingDirectory"
    }

    if (-not ('QBitLauncher.ShortcutWriter' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace QBitLauncher
{
    public static class ShortcutWriter
    {
        [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
        private class ShellLink { }

        // Keep the complete native vtable order, including unused members.
        [ComImport, Guid("000214F9-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellLinkW
        {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder path, int length, IntPtr findData, uint flags);
            void GetIDList(out IntPtr itemList);
            void SetIDList(IntPtr itemList);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder description, int length);
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string description);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder directory, int length);
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string directory);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder arguments, int length);
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string arguments);
            void GetHotkey(out short hotkey);
            void SetHotkey(short hotkey);
            void GetShowCmd(out int showCommand);
            void SetShowCmd(int showCommand);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder path, int length, out int index);
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string path, int index);
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string path, uint reserved);
            void Resolve(IntPtr window, uint flags);
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string path);
        }

        public static void Save(string shortcutPath, string targetPath, string workingDirectory, string arguments)
        {
            object shellLink = new ShellLink();
            try
            {
                IShellLinkW link = (IShellLinkW)shellLink;
                link.SetPath(targetPath);
                link.SetWorkingDirectory(workingDirectory);
                link.SetArguments(arguments ?? String.Empty);
                ((IPersistFile)shellLink).Save(shortcutPath, true);
            }
            finally
            {
                Marshal.FinalReleaseComObject(shellLink);
            }
        }
    }
}
'@ -ErrorAction Stop
    }

    [QBitLauncher.ShortcutWriter]::Save([IO.Path]::GetFullPath($ShortcutPath), $target.FullName, $WorkingDirectory, $Arguments)
}
