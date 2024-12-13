if (args.Length < 2)
{
    Console.WriteLine("Usage: MsixPacker <output.msix> <input_dir_or_file1>");
    return 1;
}

try
{
    using (var packer = new MsixPacker.MsixPacker(args[0], Path.GetFullPath(args[1])))
    {
        packer.AddFile(args[1]);
    }
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}
