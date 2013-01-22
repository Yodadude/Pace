
namespace Pace
{
	public static class Extensions
	{
		public static string Left(this string inString, int leftChars)
		{
			if (inString.Length <= leftChars) return inString;

			return inString.Substring(0, leftChars);

		}
	}
}
