using System;
using System.Collections.Generic;
using System.Diagnostics;
using XlsxSharp.Excel.Formatting;

namespace XlsxSharp.Excel.RichText;

/// <summary>
/// A class for holding <see cref="XLRichText"/> in a <see cref="SharedStringTable"/>.
/// It's immutable (keys in reverse dictionary can't change) and more memory efficient
/// than mutable rich text.
/// </summary>
[DebuggerDisplay("{Text}")]
internal sealed class XLImmutableRichText : IEquatable<XLImmutableRichText>
{
    private readonly RichTextRun[] _runs;
    private readonly PhoneticRun[] _phoneticRuns;

    private XLImmutableRichText(
        string text,
        RichTextRun[] runs,
        PhoneticRun[] phoneticRuns,
        PhoneticProperties? phoneticsProps
    )
    {
        this.Text = text;
        this._runs = runs;
        this._phoneticRuns = phoneticRuns;
        this.PhoneticsProperties = phoneticsProps;
    }

    /// <summary>
    /// A text of a whole rich text, without styling.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Individual rich text runs that make up the <see cref="Text"/>, in ascending order, non-overlapping.
    /// </summary>
    public IReadOnlyList<RichTextRun> Runs => this._runs;

    /// <summary>
    /// All phonetics runs of rich text. Empty array, if no phonetic run. In ascending order, non-overlapping.
    /// </summary>
    public IReadOnlyList<PhoneticRun> PhoneticRuns => this._phoneticRuns;

    /// <summary>
    /// Properties used to display phonetic runs.
    /// </summary>
    public PhoneticProperties? PhoneticsProperties { get; }

    public bool Equals(XLImmutableRichText? other)
    {
        if (other is null)
        {
            return false;
        }

        if (ReferenceEquals(this, other))
        {
            return true;
        }

        return this.Text == other.Text
            && this._runs.SequenceEqual(other._runs)
            && this._phoneticRuns.SequenceEqual(other._phoneticRuns)
            && Nullable.Equals(this.PhoneticsProperties, other.PhoneticsProperties);
    }

    public override bool Equals(object? obj) => this.Equals(obj as XLImmutableRichText);

    public override int GetHashCode()
    {
        HashCode hashCode = new();
        hashCode.Add(this.Text);
        hashCode.Add(this.PhoneticsProperties);
        foreach (PhoneticRun phoneticRun in this._phoneticRuns)
        {
            hashCode.Add(phoneticRun);
        }

        foreach (RichTextRun run in this._runs)
        {
            hashCode.Add(run);
        }

        return hashCode.ToHashCode();
    }

    internal string GetRunText(RichTextRun run) => this.Text.Substring(run.StartIndex, run.Length);

    /// <summary>
    /// Create an immutable rich text with same content as the original <paramref name="formattedText"/>.
    /// </summary>
    internal static XLImmutableRichText Create<T>(XLFormattedText<T> formattedText)
    {
        string text = formattedText.Text;
        RichTextRun[] runs = new RichTextRun[formattedText.Count];
        int runIdx = 0;
        int charStartIdx = 0;
        foreach (XLRichString richString in formattedText)
        {
            runs[runIdx++] = new RichTextRun(richString.Font, charStartIdx, richString.Text.Length);
            charStartIdx += richString.Text.Length;
        }

        PhoneticRun[] phoneticRuns;
        PhoneticProperties? phoneticProps;
        if (formattedText.HasPhonetics)
        {
            XLPhonetics rtPhonetics = formattedText.Phonetics;
            phoneticRuns = new PhoneticRun[rtPhonetics.Count];
            int phoneticRunIdx = 0;
            int prevPhoneticEndIdx = 0;
            foreach (IXLPhonetic phonetic in formattedText.Phonetics)
            {
                if (phonetic.Start >= text.Length)
                {
                    throw new ArgumentException(
                        "Phonetic run start index must be within the text boundaries."
                    );
                }

                if (phonetic.End > text.Length)
                {
                    throw new ArgumentException(
                        "Phonetic run end index must be at most length of a text."
                    );
                }

                if (phonetic.Start < prevPhoneticEndIdx)
                {
                    throw new ArgumentException(
                        "Phonetic runs must be in ascending order and can't overlap."
                    );
                }

                phoneticRuns[phoneticRunIdx++] = new PhoneticRun(
                    phonetic.Text,
                    phonetic.Start,
                    phonetic.End
                );
                prevPhoneticEndIdx = phonetic.End;
            }

            phoneticProps = new PhoneticProperties(formattedText.Phonetics);
        }
        else
        {
            phoneticRuns = [];
            phoneticProps = null;
        }

        return new XLImmutableRichText(text, runs, phoneticRuns, phoneticProps);
    }

    internal readonly struct RichTextRun : IEquatable<RichTextRun>
    {
        internal readonly int StartIndex;
        internal readonly int Length;
        internal readonly XLFontFormatValue Font;

        internal RichTextRun(XLFontFormatValue font, int startIndex, int length)
        {
            this.Font = font;
            this.StartIndex = startIndex;
            this.Length = length;
        }

        public bool Equals(RichTextRun other) =>
            this.StartIndex == other.StartIndex
            && this.Length == other.Length
            && this.Font.Equals(other.Font);

        public override bool Equals(object? obj) => obj is RichTextRun other && this.Equals(other);

        public override int GetHashCode() =>
            HashCode.Combine(this.StartIndex, this.Length, this.Font);
    }

    /// <summary>
    /// Phonetic runs can't overlap and must be in order (i.e. start index must be ascending).
    /// </summary>
    internal readonly struct PhoneticRun
    {
        /// <summary>
        /// Text that is displayed above a segment indicating how should segment be read.
        /// </summary>
        internal readonly string Text;

        /// <summary>
        /// Starting index of displayed phonetic (first character is 0).
        /// </summary>
        internal readonly int StartIndex;

        /// <summary>
        /// End index, excluding (the last index is a length of the rich text).
        /// </summary>
        internal readonly int EndIndex;

        public PhoneticRun(string text, int startIndex, int endIndex)
        {
            if (text.Length == 0)
            {
                throw new ArgumentException("Phonetic run text can't be empty.", nameof(text));
            }

            if (startIndex < 0)
            {
                throw new ArgumentException(
                    "Start index index must be greater than 0.",
                    nameof(startIndex)
                );
            }

            if (startIndex >= endIndex)
            {
                throw new ArgumentException(
                    "Start index must be less than end index.",
                    nameof(endIndex)
                );
            }

            this.Text = text;
            this.StartIndex = startIndex;
            this.EndIndex = endIndex;
        }

        public bool Equals(PhoneticRun other) =>
            this.Text == other.Text
            && this.StartIndex == other.StartIndex
            && this.EndIndex == other.EndIndex;

        public override bool Equals(object? obj) => obj is PhoneticRun other && this.Equals(other);

        public override int GetHashCode() =>
            HashCode.Combine(this.Text, this.StartIndex, this.EndIndex);
    }

    /// <summary>
    /// Properties of phonetic runs. All phonetic runs of a rich text have same font and other properties.
    /// </summary>
    internal readonly struct PhoneticProperties
    {
        /// <summary>
        /// Font used for text of phonetic runs. All phonetic runs use same font. There can be no phonetic runs,
        /// but with specified font (e.g. the mutable API has only specified font, but no text yet).
        /// </summary>
        public readonly XLFontFormatValue Font;

        /// <summary>
        /// Type of phonetics. Default is <see cref="XLPhoneticType.FullWidthKatakana"/>
        /// </summary>
        public readonly XLPhoneticType Type;

        /// <summary>
        /// Alignment of phonetics. Default is <see cref="XLPhoneticAlignment.Left"/>
        /// </summary>
        public readonly XLPhoneticAlignment Alignment;

        internal PhoneticProperties(XLPhonetics rtPhonetics)
        {
            this.Font = rtPhonetics.Font;
            this.Type = rtPhonetics.Type;
            this.Alignment = rtPhonetics.Alignment;
        }

        public bool Equals(PhoneticProperties other) =>
            this.Font.Equals(other.Font)
            && this.Type == other.Type
            && this.Alignment == other.Alignment;

        public override bool Equals(object? obj) =>
            obj is PhoneticProperties other && this.Equals(other);

        public override int GetHashCode() =>
            HashCode.Combine(this.Font, this.Type, this.Alignment);
    }
}
