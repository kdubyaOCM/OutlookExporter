using System.Text;

namespace OutlookExportTool.Core.Utilities;

public static class Base32
{
    // RFC 4648 Base32 alphabet (without padding)
    private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567";
    private const int BitsPerCharacter = 5;
    private const int BitsPerByte = 8;

    public static string Encode(ReadOnlySpan<byte> data)
    {
        if (data.IsEmpty)
        {
            return string.Empty;
        }

        var output = new StringBuilder((int)Math.Ceiling(data.Length * BitsPerByte / (double)BitsPerCharacter));
        int buffer = data[0];
        int next = 1;
        int bitsLeft = BitsPerByte;

        while (bitsLeft > 0 || next < data.Length)
        {
            if (bitsLeft < BitsPerCharacter)
            {
                if (next < data.Length)
                {
                    buffer <<= BitsPerByte;
                    buffer |= data[next++] & 0xff;
                    bitsLeft += BitsPerByte;
                }
                else
                {
                    int pad = BitsPerCharacter - bitsLeft;
                    buffer <<= pad;
                    bitsLeft += pad;
                }
            }

            int index = (buffer >> (bitsLeft - BitsPerCharacter)) & 0x1f;
            bitsLeft -= BitsPerCharacter;
            output.Append(Alphabet[index]);
        }

        return output.ToString();
    }
}
