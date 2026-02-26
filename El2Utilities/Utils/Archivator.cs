using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace El2Core.Utils
{

    public static class Archivator
    {
        
        private static bool _isChanged = false;
        public static bool IsChanged
        {
            get { return _isChanged; }
            set { _isChanged = value; }
        }
        private static List<ArchivatorRule> _ArchivRules = [];
        public static List<ArchivatorRule> ArchiveRules
        {
            get { return _ArchivRules; }
            set { _ArchivRules = value; _isChanged = true; }
        }
        private static string[]? _FileExtensions;
        public static string[]? FileExtensions
        {
            get { return _FileExtensions; }
            set
            {
                _isChanged = true;
                _FileExtensions = value;
            }
        }
        private static int _DelayDays =0;
        public static int DelayDays
        { 
            get { return _DelayDays; }
            set { _isChanged = true; _DelayDays = value; }
        }

        public static async Task<ArchivatorResult> ArchivateAsync(DirectoryInfo SourceLocation, int rule, CancellationToken cancellationToken = default)
        {
            
            ArchivState state = 0;
            string Location = string.Empty;
            int MovedFiles = 0;
            if (rule < 0 || rule >= ArchiveRules.Count) return new ArchivatorResult() { State = ArchivState.NoRule, Location = Location, MovedFiles = 0 };
            List<FileInfo> files;
   
            if (!SourceLocation.Exists) return new ArchivatorResult() { State = ArchivState.NoDirectory, Location = Location, MovedFiles = 0 };
            if (FileExtensions == null)
            {
                files = [.. SourceLocation.GetFiles( "*.*", SearchOption.AllDirectories)];
            }
            else
            {
                files = [];
                foreach (var ext in FileExtensions)
                {
                    files.AddRange(SourceLocation.GetFiles($"*{ext}", SearchOption.AllDirectories));
                }
            }
            if (files.Count == 0) return new ArchivatorResult() { State = ArchivState.NoFiles, Location = Location, MovedFiles = 0 };
            DirectoryInfo? arch = null;
            Location = ArchiveRules[rule].TargetPath ?? string.Empty;
            List<Task> movingtasks = [];
            foreach (var file in files)
            {

                try
                {
                    var target = file.FullName.Replace(SourceLocation.Parent.FullName, Location);
                    AdvanceFileOperations.MoveFileAsyncStream(file.FullName, Path.Combine(Location, target));
                    //movingtasks.Add(t);

                }
                catch (AggregateException ex) { throw; }
                catch (ArgumentException ex) { throw; }
                catch (FileNotFoundException ex) { throw; }
                catch (IOException ex) { throw; }
                catch (UnauthorizedAccessException ex) { throw; }
                catch (Exception ex)
                {
                    throw;
                }
            }
            //if (files.Count > 0)
            //{
            //    arch = Directory.CreateDirectory(Path.Combine(Location, dir.Name));

            //    _ = MoveFilesAsync([.. files], arch.FullName, 0);
            //}
            //foreach (var d in dir.GetDirectories())
            //{

            //    files.Clear();
            //    if (FileExtensions == null)
            //    {
            //        files.AddRange([.. d.GetFiles()]);
            //    }
            //    else
            //    {
            //        foreach (var ext in FileExtensions)
            //        {
            //            files.AddRange(d.GetFiles($"*{ext}"));
            //        }
            //    }
            //    if (files.Count > 0)
            //    {
            //        if (arch == null || !arch.Exists)
            //        {
            //            arch = Directory.CreateDirectory(Path.Combine(Location, dir.Name));
            //        }

            //        var subArch = Directory.CreateDirectory(Path.Combine(arch.FullName, d.Name));
            //        ValueTuple<int, int> result = new (0,0);
            //        do
            //        {
            //           result = MoveFilesAsync([.. files], subArch.FullName, result.Item2).Result;
            //           MovedFiles += result.Item1;
            //        } while (result.Item2 > 0);
            //    }
            //}
            if (movingtasks.Count > 0)
            {
                Task.WaitAll(movingtasks);
            }
            MovedFiles = files.Count;
            if (MovedFiles > 0)
            {  
                state = ArchivState.Archivated;
            }
            return new ArchivatorResult() { State = state, Location = Location, MovedFiles = MovedFiles };
        }
        private static void MoveFiles(FileInfo[] source, string target, ref ArchivState state)
        {
            foreach (var file in source)
            {
                try
                {
                    file.MoveTo(Path.Combine(target, file.Name));
                    state = ArchivState.Archivated;
                }
                catch (Exception ex)
                {
                    state = 0;
                }
            }
        }
        private static async Task<ValueTuple<int, string>> MoveFilesAsync(FileInfo[] source, string target, int repeatNr)
        {
            int result = 0;
            int dirCount = 0;
            string path = string.Empty;
            await Task.Run(() =>
            {

                foreach (var file in source)
                {
  
                    try
                    {
                        AdvanceFileOperations.MoveFileAsync(file.FullName, Path.Combine(target, file.Name)).Wait();
                        result++;
                    }
                    catch (Exception ex)
                    {
                        throw;
                    }
                    dirCount = file.Directory != null ? file.Directory.GetDirectories().Length : 0;
                    if (dirCount > 0) path = file.Directory.GetDirectories()[repeatNr].FullName;
                }            
            });
            return new (result, path);
        }

        public enum ArchivatorTarget
        {
            TTNR,
            OrderNumber
        }
        public enum ArchivState
        {
            None,
            Archivated,
            NoFiles,
            NoDirectory,
            NoRule
        }
    }
    public struct ArchivatorResult
    {
        public Archivator.ArchivState State { get; set; }
        public string Location { get; set; }
        public int MovedFiles { get; set; }
    }
    public class ArchivatorRule
    {
        public string? Name { get; set; }
        public string? RegexString { get; set; }
        public Archivator.ArchivatorTarget MatchTarget { get; set; } = Archivator.ArchivatorTarget.TTNR;
        public string? TargetPath { get; set; }
        public ArchivatorRule(string name, string regex, Archivator.ArchivatorTarget target, string targetPath)
        {
            Name = name;
            RegexString = regex;
            MatchTarget = target;
            TargetPath  = targetPath;
  
        }
        public ArchivatorRule() { }
    }   
}
