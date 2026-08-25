namespace XlsxSharp.ExcelNumberFormat;

internal class Tokenizer
{
    private readonly ReadOnlyMemory<char> formatString;
    private int formatStringPosition;

    public Tokenizer(string fmt) => this.formatString = fmt.AsMemory();

    public Tokenizer(ReadOnlyMemory<char> fmt) => this.formatString = fmt;

    public int Position => this.formatStringPosition;

    public int Length => this.formatString.Length;

    public ReadOnlyMemory<char> Slice(int startIndex, int length) =>
        this.formatString.Slice(startIndex, length);

    public int Peek(int offset = 0)
    {
        if (this.formatStringPosition + offset >= this.Length)
        {
            return -1;
        }

        return this.formatString.Span[this.formatStringPosition + offset];
    }

    public int PeekUntil(int startOffset, int until)
    {
        int offset = startOffset;
        while (true)
        {
            int c = this.Peek(offset++);
            if (c == -1)
            {
                break;
            }

            if (c == until)
            {
                return offset - startOffset;
            }
        }
        return 0;
    }

    public bool PeekOneOf(int offset, string s)
    {
        foreach (char c in s)
        {
            if (this.Peek(offset) == c)
            {
                return true;
            }
        }
        return false;
    }

    public void Advance(int characters = 1) =>
        this.formatStringPosition = Math.Min(
            this.formatStringPosition + characters,
            this.formatString.Length
        );

    public bool ReadOneOrMore(int c)
    {
        if (this.Peek() != c)
        {
            return false;
        }

        while (this.Peek() == c)
        {
            this.Advance();
        }

        return true;
    }

    public bool ReadOneOf(string s)
    {
        if (this.PeekOneOf(0, s))
        {
            this.Advance();
            return true;
        }
        return false;
    }

    public bool ReadString(string s, bool ignoreCase = false)
    {
        if (this.formatStringPosition + s.Length > this.Length)
        {
            return false;
        }

        for (int i = 0; i < s.Length; i++)
        {
            char c1 = s[i];
            char c2 = (char)this.Peek(i);
            if (ignoreCase)
            {
                if (char.ToLower(c1) != char.ToLower(c2))
                {
                    return false;
                }
            }
            else
            {
                if (c1 != c2)
                {
                    return false;
                }
            }
        }

        this.Advance(s.Length);
        return true;
    }

    public bool ReadEnclosed(char open, char close)
    {
        if (this.Peek() == open)
        {
            int length = this.PeekUntil(1, close);
            if (length > 0)
            {
                this.Advance(1 + length);
                return true;
            }
        }

        return false;
    }
}
