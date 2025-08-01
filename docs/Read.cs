using Frends.Files.Read.Definitions;
using Microsoft.Win32.SafeHandles;
using SimpleImpersonation;
using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace Frends.Files.Read;

///<summary>
/// Files task.
/// </summary>
public class Files
{
    /// <summary>
    /// Reads a file from directory.
    /// [Documentation](https://tasks.frends.com/tasks/frends-tasks/Frends.Files.Read)
    /// </summary>
    /// <param name="input">Input parameters</param>
    /// <param name="options">Options parameters</param>
    /// <returns>Object { string Content, string Path, double SizeInMegaBytes, DateTime CreationTime, DateTime LastWriteTime }</returns>
    public static async Task<Result> Read([PropertyTab] Input input, [PropertyTab] Options options)
    {
        return await ExecuteActionAsync(
                    () => ExecuteRead(input, options),
                    options.UseGivenUserCredentialsForRemoteConnections,
                    options.UserName,
                    options.Password)
                .ConfigureAwait(false);
    }

    private static async Task<TResult> ExecuteActionAsync<TResult>(Func<Task<TResult>> action, bool useGivenCredentials, string username, string password)
    {
        if (!useGivenCredentials)
            return await action().ConfigureAwait(false);

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            throw new PlatformNotSupportedException("UseGivenCredentials feature is only supported on Windows.");

        var (domain, user) = GetDomainAndUsername(username);

        UserCredentials credentials = new(domain, user, password);
        using SafeAccessTokenHandle userHandle = credentials.LogonUser(LogonType.NewCredentials);

        return await WindowsIdentity.RunImpersonated(userHandle, async () => await action().ConfigureAwait(false));
    }

    private static async Task<Result> ExecuteRead(Input input, Options options)
    {
        Encoding encoding = GetEncoding(options.FileEncoding, options.EnableBom, options.EncodingInString);

        using var fileStream = new FileStream(input.Path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, useAsync: true);
        using var reader = new StreamReader(fileStream, encoding, detectEncodingFromByteOrderMarks: true);
        var content = await reader.ReadToEndAsync().ConfigureAwait(false);
        return new Result(new FileInfo(input.Path), content);
    }

    internal static Tuple<string, string> GetDomainAndUsername(string username)
    {
        var domainAndUserName = username.Split('\\');
        if (domainAndUserName.Length != 2)
            throw new ArgumentException($@"UserName field must be of format domain\username was: {username}");
        return new Tuple<string, string>(domainAndUserName[0], domainAndUserName[1]);
    }

    private static Encoding GetEncoding(FileEncoding optionsFileEncoding, bool optionsEnableBom, string optionsEncodingInString)
    {
        switch (optionsFileEncoding)
        {
            case FileEncoding.Other:
                return Encoding.GetEncoding(optionsEncodingInString);
            case FileEncoding.ASCII:
                return Encoding.ASCII;
            case FileEncoding.Default:
                return Encoding.Default;
            case FileEncoding.UTF8:
                return optionsEnableBom ? new UTF8Encoding(true) : new UTF8Encoding(false);
            case FileEncoding.Windows1252:
                EncodingProvider provider = CodePagesEncodingProvider.Instance;
                Encoding.RegisterProvider(provider);
                return Encoding.GetEncoding(1252);
            case FileEncoding.Unicode:
                return Encoding.Unicode;
            default:
                throw new ArgumentOutOfRangeException();
        }
    }
}