namespace Sheetwright;

/// <summary>
/// Chooses a CSV file's separator from its first rows.
/// </summary>
/// <remarks>
/// <para>
/// Excel writes the separator of the machine's regional settings, so a file from a Turkish or German machine
/// is separated by semicolons and one from an English machine by commas; a person who copies a table out of a
/// browser has tabs. There is no marker in the file (unless <c>sep=</c> is there), so the choice is a measure of
/// consistency: for each candidate the first rows are split the way a reader would split them (a separator
/// inside quotes does not count), and the candidate that gives the most rows the same number of fields, above
/// one, wins. Ties go to comma, semicolon, tab, pipe in that order. A file in which nothing splits (one column)
/// is read as comma-separated.
/// </para>
/// <para>
/// Only the sample is looked at, so a file whose first rows are not typical is not detected; the caller can
/// always set <see cref="SpreadsheetReaderOptions.Delimiter"/>.
/// </para>
/// </remarks>
internal static class SpreadsheetCsvDelimiter
{
    private const int SampleRecords = 50;

    private static readonly char[] Candidates = [',', ';', '\t', '|'];

    /// <summary>
    /// Whether the text starts with an Excel <c>sep=</c> line, and the character it names.
    /// </summary>
    internal static bool TryReadDirective(string sample, out char separator)
    {
        separator = ',';
        if (sample.Length < 6
            || !sample.StartsWith("sep=", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        char named = sample[4];
        char after = sample[5];
        if (named is '"' or '\r' or '\n' || (after != '\r' && after != '\n'))
        {
            return false;
        }

        separator = named;
        return true;
    }

    /// <summary>The most consistent candidate separator for the sample.</summary>
    /// <param name="sample">The decoded start of the file.</param>
    /// <param name="truncated">Whether the sample ends before the file does, so its last row may be cut.</param>
    internal static char Detect(string sample, bool truncated)
    {
        char best = ',';
        int bestAgree = 0;
        int bestFields = 0;

        foreach (char candidate in Candidates)
        {
            List<int> counts = CountFields(sample, candidate, truncated);
            if (counts.Count == 0)
            {
                continue;
            }

            int mode = counts.GroupBy(count => count).OrderByDescending(group => group.Count()).ThenByDescending(group => group.Key).First().Key;
            if (mode < 2)
            {
                continue;
            }

            int agree = counts.Count(count => count == mode);
            if (agree > bestAgree || (agree == bestAgree && mode > bestFields))
            {
                best = candidate;
                bestAgree = agree;
                bestFields = mode;
            }
        }

        return best;
    }

    private static List<int> CountFields(string sample, char delimiter, bool truncated)
    {
        List<int> counts = [];
        int fields = 1;
        bool inQuotes = false;
        bool fieldStart = true;
        bool recordHasContent = false;

        for (int index = 0; index < sample.Length && counts.Count < SampleRecords; index++)
        {
            char character = sample[index];

            if (inQuotes)
            {
                if (character == '"')
                {
                    if (index + 1 < sample.Length && sample[index + 1] == '"')
                    {
                        index++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }

                continue;
            }

            if (character == '"' && fieldStart)
            {
                inQuotes = true;
                fieldStart = false;
                recordHasContent = true;
            }
            else if (character == delimiter)
            {
                fields++;
                fieldStart = true;
                recordHasContent = true;
            }
            else if (character is '\r' or '\n')
            {
                if (character == '\r' && index + 1 < sample.Length && sample[index + 1] == '\n')
                {
                    index++;
                }

                if (recordHasContent)
                {
                    counts.Add(fields);
                }

                fields = 1;
                fieldStart = true;
                recordHasContent = false;
            }
            else
            {
                fieldStart = false;
                recordHasContent = true;
            }
        }

        // A last row with no line break is a whole row when the file ends there and half of one when the sample
        // was cut, and half of one says nothing about the file.
        if (!truncated && !inQuotes && recordHasContent && counts.Count < SampleRecords)
        {
            counts.Add(fields);
        }

        return counts;
    }
}
