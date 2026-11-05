namespace ExcelFormulaRemover.Core;

using System.Buffers.Binary;
using System.Text;

/// <summary>
/// One [MS-XLSB] 2.1.4 record: a variable-length type and size header followed by its payload.
/// </summary>
internal readonly record struct Biff12Record(int Type, byte[] Data)
{
    /// <summary>The largest size that the four-byte, 7-bit size header can express.</summary>
    public const int MaximumDataLength = (1 << 28) - 1;

    /// <summary>The largest type that the two-byte, 7-bit type header can express.</summary>
    public const int MaximumType = (1 << 14) - 1;

    public uint ReadUInt32(int offset)
    {
        RequireLength(offset + sizeof(uint));
        return BinaryPrimitives.ReadUInt32LittleEndian(Data.AsSpan(offset));
    }

    /// <summary>Reads an XLWideString or XLNullableWideString; returns null for the nullable marker.</summary>
    public string? ReadWideString(int offset, out int nextOffset)
    {
        uint characterCount = ReadUInt32(offset);
        if (characterCount == uint.MaxValue)
        {
            nextOffset = offset + sizeof(uint);
            return null;
        }

        long byteCount = characterCount * 2L;
        if (byteCount > Data.Length - offset - sizeof(uint))
        {
            throw new InvalidDataException($"Binary workbook record {Type} has an invalid string length.");
        }

        nextOffset = offset + sizeof(uint) + (int)byteCount;
        return Encoding.Unicode.GetString(Data, offset + sizeof(uint), (int)byteCount);
    }

    public void RequireLength(int length)
    {
        if (Data.Length < length)
        {
            throw new InvalidDataException(
                $"Binary workbook record {Type} is {Data.Length} bytes; at least {length} are required.");
        }
    }
}

/// <summary>
/// Reads a forward-only sequence of BIFF12 records and rejects truncated framing.
/// </summary>
internal sealed class Biff12Reader(Stream stream)
{
    private const int DirectReadLength = 1024 * 1024;

    public bool TryRead(out Biff12Record record)
    {
        int first = stream.ReadByte();
        if (first < 0)
        {
            record = default;
            return false;
        }

        int type = first & 0x7F;
        if ((first & 0x80) != 0)
        {
            int second = ReadRequiredByte();
            if ((second & 0x80) != 0)
            {
                throw new InvalidDataException("A binary workbook record type is longer than two bytes.");
            }

            type |= second << 7;
        }

        int length = 0;
        for (int index = 0; ; index++)
        {
            int value = ReadRequiredByte();
            length |= (value & 0x7F) << (7 * index);
            if ((value & 0x80) == 0)
            {
                break;
            }

            if (index == 3)
            {
                throw new InvalidDataException("A binary workbook record size is longer than four bytes.");
            }
        }

        record = new Biff12Record(type, ReadData(length));
        return true;
    }

    /// <summary>
    /// Reads a record payload. Payloads above <see cref="DirectReadLength"/> grow with the bytes
    /// actually present, so a corrupt size header fails at end of stream instead of first
    /// allocating up to 256 MB.
    /// </summary>
    private byte[] ReadData(int length)
    {
        try
        {
            if (length <= DirectReadLength)
            {
                byte[] data = new byte[length];
                stream.ReadExactly(data);
                return data;
            }

            using MemoryStream buffer = new(DirectReadLength);
            byte[] chunk = new byte[DirectReadLength];
            int remaining = length;
            while (remaining > 0)
            {
                int count = Math.Min(remaining, chunk.Length);
                stream.ReadExactly(chunk, 0, count);
                buffer.Write(chunk, 0, count);
                remaining -= count;
            }

            return buffer.ToArray();
        }
        catch (EndOfStreamException exception)
        {
            throw new InvalidDataException("A binary workbook record is truncated.", exception);
        }
    }

    private int ReadRequiredByte()
    {
        int value = stream.ReadByte();
        return value >= 0
            ? value
            : throw new InvalidDataException("A binary workbook record header is truncated.");
    }
}

/// <summary>
/// Writes BIFF12 records with the minimal header encoding required by [MS-XLSB] 2.1.4.
/// </summary>
internal static class Biff12Writer
{
    public static void Write(Stream stream, int type, ReadOnlySpan<byte> data)
    {
        if (type is < 0 or > Biff12Record.MaximumType)
        {
            throw new ArgumentOutOfRangeException(nameof(type), type, "Invalid BIFF12 record type.");
        }

        if (data.Length > Biff12Record.MaximumDataLength)
        {
            throw new InvalidDataException("A binary workbook record is too large.");
        }

        Span<byte> header = stackalloc byte[6];
        int headerLength = 0;
        if (type < 0x80)
        {
            header[headerLength++] = (byte)type;
        }
        else
        {
            header[headerLength++] = (byte)((type & 0x7F) | 0x80);
            header[headerLength++] = (byte)(type >> 7);
        }

        int remaining = data.Length;
        do
        {
            byte value = (byte)(remaining & 0x7F);
            remaining >>= 7;
            if (remaining != 0)
            {
                value |= 0x80;
            }

            header[headerLength++] = value;
        }
        while (remaining != 0);

        stream.Write(header[..headerLength]);
        stream.Write(data);
    }

    public static void Write(Stream stream, Biff12Record record) =>
        Write(stream, record.Type, record.Data);
}
