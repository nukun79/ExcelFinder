using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace ExcelFinder;

public static class PerforceHelper
{
    private static readonly Regex OpenedByClientRegex = new(
        @"\bby\s+[^@\s]+@(?<client>[^\s]+)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static readonly Regex FilelogLineRegex = new(
        @"^\.\.\.\s+(?:\.\.\.\s+)?#(?<rev>\d+)\s+change\s+(?<change>\d+)\s+(?<action>\S+)\s+on\s+(?<date>.+?)\s+by\s+(?<user>[^@]+)@(?<client>\S+)\s+\((?<filetype>[^)]+)\)(?:\s+'(?<desc>.*)')?$",
        RegexOptions.Compiled);

    public static (bool success, string infoText) GetInfoRaw()
    {
        return RunP4(["info"]);
    }

    public static (bool success, string message, PerforceClientSettings? settings) GetClientSettings()
    {
        (bool success, string message) = GetInfoRaw();
        if (!success)
        {
            return (false, message, null);
        }

        string clientName = ReadValue(message, "Client name:");
        string host = ReadValue(message, "Client host:");
        string root = ReadValue(message, "Client root:");
        string stream = ReadValue(message, "Client stream:");

        return (true, "성공", new PerforceClientSettings
        {
            Client = clientName,
            Host = host,
            Root = root,
            Stream = stream
        });
    }

    public static (bool success, string message) Checkout(string filePath)
    {
        return RunP4(["edit", filePath]);
    }

    public static (bool success, string message) Revert(string filePath)
    {
        return RunP4(["revert", filePath]);
    }

    public static (bool success, string message) Checkin(string filePath, string description)
    {
        string submitDescription = string.IsNullOrWhiteSpace(description)
            ? $"ExcelFinder submit: {Path.GetFileName(filePath)}"
            : description.Trim();

        return RunP4(["submit", "-d", submitDescription, filePath]);
    }

    public static (bool success, bool isCheckedOut, string message) GetCheckoutStatus(string filePath)
    {
        (bool success, string message) = RunP4(["opened", filePath]);
        string lowered = message.ToLowerInvariant();
        bool notOpened = lowered.Contains("not opened") || lowered.Contains("file(s) not opened");
        if (notOpened)
        {
            return (true, false, "체크아웃 안됨");
        }

        bool openedPattern =
            lowered.Contains(" - edit") ||
            lowered.Contains(" - add") ||
            lowered.Contains(" - delete") ||
            lowered.Contains(" - branch") ||
            lowered.Contains(" - integrate") ||
            lowered.Contains(" - move/add") ||
            lowered.Contains(" - move/delete") ||
            lowered.Contains(" - import");

        if (openedPattern)
        {
            return (true, true, "체크아웃됨");
        }

        if (success)
        {
            return (true, false, "체크아웃 안됨");
        }

        return (false, false, message);
    }

    public static (bool success, string message, List<string> clients) GetOpenedClientWorkspaces(string filePath)
    {
        (bool success, string message) = RunP4(["opened", "-a", filePath]);
        if (!success)
        {
            return (false, message, []);
        }

        var clients = new List<string>();
        foreach (string rawLine in message.Replace("\r\n", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            string line = rawLine.Trim();
            Match m = OpenedByClientRegex.Match(line);
            if (!m.Success)
            {
                continue;
            }

            string client = m.Groups["client"].Value.Trim();
            if (string.IsNullOrWhiteSpace(client))
            {
                continue;
            }

            if (!clients.Contains(client, StringComparer.OrdinalIgnoreCase))
            {
                clients.Add(client);
            }
        }

        return (true, "성공", clients);
    }

    public static (bool success, string message, List<PerforceFileHistoryEntry> entries) GetFileHistory(string filePath)
    {
        (bool success, string message) = RunP4(["filelog", "-t", "-l", filePath]);
        if (!success)
        {
            return (false, message, []);
        }

        List<PerforceFileHistoryEntry> entries = ParseFilelogEntries(message);
        if (entries.Count > 0)
        {
            return (true, "성공", entries);
        }

        // 로컬 경로 전달 시 파싱이 비거나 파일로그가 얕게 나오는 경우 depot 경로로 재시도한다.
        (bool whereOk, string whereMsg) = RunP4(["where", filePath]);
        if (!whereOk)
        {
            return (true, "성공", entries);
        }

        string? depotPath = ExtractDepotPathFromWhere(whereMsg);
        if (string.IsNullOrWhiteSpace(depotPath))
        {
            return (true, "성공", entries);
        }

        (bool retryOk, string retryMsg) = RunP4(["filelog", "-t", "-l", depotPath]);
        if (!retryOk)
        {
            return (true, "성공", entries);
        }

        entries = ParseFilelogEntries(retryMsg);
        return (true, "성공", entries);
    }

    public static (bool success, string message, string depotPath) GetDepotPathForLocalFile(string localFilePath)
    {
        (bool success, string message) = RunP4(["where", localFilePath]);
        if (!success)
        {
            return (false, message, string.Empty);
        }

        string? depotPath = ExtractDepotPathFromWhere(message);
        if (string.IsNullOrWhiteSpace(depotPath))
        {
            return (false, "depot 경로를 찾을 수 없습니다.", string.Empty);
        }

        return (true, "성공", depotPath);
    }

    public static (bool success, string message) ExportDepotRevisionToFile(string depotFileWithRevision, string outputPath)
    {
        return RunP4(["print", "-q", "-o", outputPath, depotFileWithRevision]);
    }

    private static List<PerforceFileHistoryEntry> ParseFilelogEntries(string message)
    {
        var entries = new List<PerforceFileHistoryEntry>();
        string[] lines = message.Replace("\r\n", "\n").Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i].TrimEnd();
            Match m = FilelogLineRegex.Match(line);
            if (!m.Success)
            {
                continue;
            }

            var descLines = new List<string>();
            string inlineDesc = m.Groups["desc"].Value;
            if (!string.IsNullOrWhiteSpace(inlineDesc))
            {
                descLines.Add(inlineDesc);
            }

            int cursor = i + 1;
            while (cursor < lines.Length)
            {
                string nextLine = lines[cursor].TrimEnd();
                if (FilelogLineRegex.IsMatch(nextLine))
                {
                    break;
                }

                string trimmedStart = nextLine.TrimStart();
                if (trimmedStart.StartsWith("//", StringComparison.Ordinal))
                {
                    break;
                }

                string? normalizedDesc = NormalizeDescriptionLine(nextLine);
                if (!string.IsNullOrWhiteSpace(normalizedDesc))
                {
                    descLines.Add(normalizedDesc);
                }

                cursor++;
            }

            i = cursor - 1;

            entries.Add(new PerforceFileHistoryEntry
            {
                Revision = m.Groups["rev"].Value,
                Changelist = m.Groups["change"].Value,
                DateSubmitted = m.Groups["date"].Value,
                SubmittedBy = m.Groups["user"].Value,
                Client = m.Groups["client"].Value,
                Action = m.Groups["action"].Value,
                FileType = m.Groups["filetype"].Value,
                Description = string.Join(Environment.NewLine, descLines).Trim()
            });
        }

        return entries;
    }

    private static string? NormalizeDescriptionLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return null;
        }

        string trimmedStart = line.TrimStart();

        const string prefix = "... ...";
        if (trimmedStart.StartsWith(prefix, StringComparison.Ordinal))
        {
            return trimmedStart[prefix.Length..].TrimStart();
        }

        // 서버/환경에 따라 -l 설명 줄이 탭/공백 들여쓰기로만 내려오는 경우를 허용한다.
        if (line.StartsWith('\t') || line.StartsWith("    ", StringComparison.Ordinal))
        {
            return trimmedStart;
        }

        return null;
    }

    private static string? ExtractDepotPathFromWhere(string whereOutput)
    {
        string[] lines = whereOutput.Replace("\r\n", "\n").Split('\n', StringSplitOptions.RemoveEmptyEntries);
        foreach (string line in lines)
        {
            string trimmed = line.Trim();
            if (trimmed.StartsWith("//", StringComparison.Ordinal))
            {
                string[] tokens = trimmed.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length > 0 && tokens[0].StartsWith("//", StringComparison.Ordinal))
                {
                    return tokens[0];
                }
            }
        }

        return null;
    }

    public static (bool success, string message) ApplyClientSettings(PerforceClientSettings settings)
    {
        if (string.IsNullOrWhiteSpace(settings.Client))
        {
            return (false, "Client 값을 입력해 주세요.");
        }

        (bool ok, string specOutput) = RunP4(["client", "-o", settings.Client]);
        if (!ok)
        {
            return (false, $"client spec 조회 실패: {specOutput}");
        }

        string updatedSpec = ReplaceSpecField(specOutput, "Client", settings.Client.Trim());
        updatedSpec = ReplaceSpecField(updatedSpec, "Host", settings.Host.Trim());
        updatedSpec = ReplaceSpecField(updatedSpec, "Root", settings.Root.Trim());
        updatedSpec = ReplaceSpecField(updatedSpec, "Stream", settings.Stream.Trim());

        (bool applyOk, string applyMessage) = RunP4(["client", "-i"], updatedSpec);
        if (!applyOk)
        {
            return (false, applyMessage);
        }

        (bool setOk, string setMessage) = RunP4(["set", $"P4CLIENT={settings.Client.Trim()}"]);
        if (!setOk)
        {
            return (false, $"client 적용 후 P4CLIENT 설정 실패: {setMessage}");
        }

        return (true, "Perforce Client 정보 적용 완료");
    }

    private static string ReadValue(string source, string key)
    {
        foreach (string line in source.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.StartsWith(key, StringComparison.OrdinalIgnoreCase))
            {
                return line.Substring(key.Length).Trim();
            }
        }

        return string.Empty;
    }

    private static string ReplaceSpecField(string spec, string key, string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return spec;
        }

        string marker = $"{key}:";
        string[] lines = spec.Replace("\r\n", "\n").Split('\n');
        bool replaced = false;

        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            if (line.StartsWith(marker, StringComparison.OrdinalIgnoreCase))
            {
                lines[i] = $"{key}: {value}";
                replaced = true;
                break;
            }
        }

        if (!replaced)
        {
            var lineList = lines.ToList();
            lineList.Insert(0, $"{key}: {value}");
            lines = lineList.ToArray();
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static (bool success, string message) RunP4(IEnumerable<string> args, string? standardInput = null)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "p4",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            RedirectStandardInput = standardInput is not null,
            CreateNoWindow = true
        };

        foreach (string arg in args)
        {
            psi.ArgumentList.Add(arg);
        }

        try
        {
            using var process = new Process { StartInfo = psi };
            process.Start();

            if (standardInput is not null)
            {
                process.StandardInput.Write(standardInput);
                process.StandardInput.Close();
            }

            string stdOut = process.StandardOutput.ReadToEnd();
            string stdErr = process.StandardError.ReadToEnd();
            process.WaitForExit();

            string output = $"{stdOut}\n{stdErr}".Trim();
            if (string.IsNullOrWhiteSpace(output))
            {
                output = process.ExitCode == 0 ? "성공" : "실패";
            }

            return (process.ExitCode == 0, output);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }
}

public sealed class PerforceClientSettings
{
    public string Client { get; set; } = string.Empty;
    public string Host { get; set; } = string.Empty;
    public string Root { get; set; } = string.Empty;
    public string Stream { get; set; } = string.Empty;
}

public sealed class PerforceFileHistoryEntry
{
    public string Revision { get; set; } = string.Empty;
    public string Changelist { get; set; } = string.Empty;
    public string DateSubmitted { get; set; } = string.Empty;
    public string SubmittedBy { get; set; } = string.Empty;
    public string Client { get; set; } = string.Empty;
    public string Action { get; set; } = string.Empty;
    public string FileType { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}
