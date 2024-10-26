if ((args?.Length) >= 2 && int.TryParse(args[1], out var size))
{
    var file = new FileInfo(args[0]);
    if (file.Exists)
    {
        Console.WriteLine("The file already exists.");
        return 2;
    }
    using var fs = file.Create();
    fs.Write(new byte[size]);
    return 0;
}
else
{
    Console.WriteLine("Usage:   allocfile filename filelength\nExample: allocfile myfile.dat 10000");
    return 1;
}