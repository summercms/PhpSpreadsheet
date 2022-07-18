<?php

namespace PhpOffice\PhpSpreadsheet\Calculation\TextData;

use PhpOffice\PhpSpreadsheet\Calculation\ArrayEnabled;
use PhpOffice\PhpSpreadsheet\Calculation\Exception as CalcExp;
use PhpOffice\PhpSpreadsheet\Calculation\Functions;
use PhpOffice\PhpSpreadsheet\Calculation\Information\ExcelError;
use PhpOffice\PhpSpreadsheet\Shared\StringHelper;

class Extract
{
    use ArrayEnabled;

    /**
     * LEFT.
     *
     * @param mixed $value String value from which to extract characters
     *                         Or can be an array of values
     * @param mixed $chars The number of characters to extract (as an integer)
     *                         Or can be an array of values
     *
     * @return array|string The joined string
     *         If an array of values is passed for the $value or $chars arguments, then the returned result
     *            will also be an array with matching dimensions
     */
    public static function left($value, $chars = 1)
    {
        if (is_array($value) || is_array($chars)) {
            return self::evaluateArrayArguments([self::class, __FUNCTION__], $value, $chars);
        }

        try {
            $value = Helpers::extractString($value);
            $chars = Helpers::extractInt($chars, 0, 1);
        } catch (CalcExp $e) {
            return $e->getMessage();
        }

        return mb_substr($value ?? '', 0, $chars, 'UTF-8');
    }

    /**
     * MID.
     *
     * @param mixed $value String value from which to extract characters
     *                         Or can be an array of values
     * @param mixed $start Integer offset of the first character that we want to extract
     *                         Or can be an array of values
     * @param mixed $chars The number of characters to extract (as an integer)
     *                         Or can be an array of values
     *
     * @return array|string The joined string
     *         If an array of values is passed for the $value, $start or $chars arguments, then the returned result
     *            will also be an array with matching dimensions
     */
    public static function mid($value, $start, $chars)
    {
        if (is_array($value) || is_array($start) || is_array($chars)) {
            return self::evaluateArrayArguments([self::class, __FUNCTION__], $value, $start, $chars);
        }

        try {
            $value = Helpers::extractString($value);
            $start = Helpers::extractInt($start, 1);
            $chars = Helpers::extractInt($chars, 0);
        } catch (CalcExp $e) {
            return $e->getMessage();
        }

        return mb_substr($value ?? '', --$start, $chars, 'UTF-8');
    }

    /**
     * RIGHT.
     *
     * @param mixed $value String value from which to extract characters
     *                         Or can be an array of values
     * @param mixed $chars The number of characters to extract (as an integer)
     *                         Or can be an array of values
     *
     * @return array|string The joined string
     *         If an array of values is passed for the $value or $chars arguments, then the returned result
     *            will also be an array with matching dimensions
     */
    public static function right($value, $chars = 1)
    {
        if (is_array($value) || is_array($chars)) {
            return self::evaluateArrayArguments([self::class, __FUNCTION__], $value, $chars);
        }

        try {
            $value = Helpers::extractString($value);
            $chars = Helpers::extractInt($chars, 0, 1);
        } catch (CalcExp $e) {
            return $e->getMessage();
        }

        return mb_substr($value ?? '', mb_strlen($value ?? '', 'UTF-8') - $chars, $chars, 'UTF-8');
    }

    /**
     * TEXTBEFORE.
     *
     * @param ?string $text the text that you're searching
     * @param ?string $delimiter the text that marks the point before which you want to extract
     * @param ?int $instance The instance of the delimiter after which you want to extract the text.
     *                      By default, this is the first instance (1).
     *                      A negative value means start searching from the end of the text string.
     * @param ?int $matchMode Determines whether the match is case-sensitive or not.
     *                       0 - Case-sensitive
     *                       1 - Case-insensitive
     * @param ?int $matchEnd Treats the end of text as a delimiter.
     *                      0 - Don't match the delimiter against the end of the text.
     *                      1 - Match the delimiter against the end of the text.
     * @param mixed $ifNotFound value to return if no match is found
     *                          The default is a #N/A Error
     *
     * @return mixed the string extracted from text before the delimiter; or the $ifNotFound value
     */
    public static function before($text, $delimiter, $instance = 1, $matchMode = 0, $matchEnd = 0, $ifNotFound = '#N/A')
    {
        $text = Helpers::extractString(Functions::flattenSingleValue($text ?? ''));
        $delimiter = Helpers::extractString(Functions::flattenSingleValue($delimiter ?? ''));
        $instance = (int) Functions::flattenSingleValue($instance ?? 1);
        $matchMode = (int) Functions::flattenSingleValue($matchMode ?? 0);
        $matchEnd = (int) Functions::flattenSingleValue($matchEnd ?? 0);
        $ifNotFound = Functions::flattenSingleValue($ifNotFound ?? ExcelError::NA());

        $split = self::validateTextBeforeAfter($text, $delimiter, $instance, $matchMode, $matchEnd, $ifNotFound);
        if (is_array($split) === false) {
            return $split;
        }
        if ($delimiter === '') {
            return ($instance > 0) ? '' : $text;
        }

        $split = ($instance < 0)
            ? array_slice($split, 0, count($split) + ($instance * 2 + 1))
            : array_slice($split, 0, $instance * 2 - 1);

        return implode('', $split);
    }

    /**
     * TEXTAFTER.
     *
     * @param ?string $text the text that you're searching
     * @param ?string $delimiter the text that marks the point before which you want to extract
     * @param ?int $instance The instance of the delimiter after which you want to extract the text.
     *                      By default, this is the first instance (1).
     *                      A negative value means start searching from the end of the text string.
     * @param ?int $matchMode Determines whether the match is case-sensitive or not.
     *                       0 - Case-sensitive
     *                       1 - Case-insensitive
     * @param ?int $matchEnd Treats the end of text as a delimiter.
     *                      0 - Don't match the delimiter against the end of the text.
     *                      1 - Match the delimiter against the end of the text.
     * @param ?mixed $ifNotFound value to return if no match is found
     *                          The default is a #N/A Error
     *
     * @return mixed the string extracted from text before the delimiter; or the $ifNotFound value
     */
    public static function after($text, $delimiter, $instance = 1, $matchMode = 0, $matchEnd = 0, $ifNotFound = '#N/A')
    {
        $text = Helpers::extractString(Functions::flattenSingleValue($text ?? ''));
        $delimiter = Helpers::extractString(Functions::flattenSingleValue($delimiter ?? ''));
        $instance = (int) Functions::flattenSingleValue($instance ?? 1);
        $matchMode = (int) Functions::flattenSingleValue($matchMode ?? 0);
        $matchEnd = (int) Functions::flattenSingleValue($matchEnd ?? 0);
        $ifNotFound = Functions::flattenSingleValue($ifNotFound ?? ExcelError::NA());

        $split = self::validateTextBeforeAfter($text, $delimiter, $instance, $matchMode, $matchEnd, $ifNotFound);
        if (is_array($split) === false) {
            return $split;
        }
        if ($delimiter === '') {
            return ($instance < 0) ? '' : $text;
        }

        $split = ($instance < 0)
            ? array_slice($split, count($split) - (abs($instance + 1) * 2))
            : array_slice($split, $instance * 2);

        return implode('', $split);
    }

    /**
     * @param int $matchMode
     * @param int $matchEnd
     * @param mixed $ifNotFound
     *
     * @return string|string[]
     */
    private static function validateTextBeforeAfter(string $text, string $delimiter, int $instance, $matchMode, $matchEnd, $ifNotFound)
    {
        $flags = ($matchMode === 0) ? 'mu' : 'miu';

        if (preg_match('/' . preg_quote($delimiter) . "/{$flags}", $text) === 0 && $matchEnd === 0) {
            return $ifNotFound;
        }

        $split = preg_split('/(' . preg_quote($delimiter) . ")/{$flags}", $text, 0, PREG_SPLIT_NO_EMPTY | PREG_SPLIT_DELIM_CAPTURE);
        if ($split === false) {
            return ExcelError::NA();
        }

        if ($instance === 0 || abs($instance) > StringHelper::countCharacters($text)) {
            return ExcelError::VALUE();
        }

        if ($matchEnd === 0 && (abs($instance) > floor(count($split) / 2))) {
            return ExcelError::NA();
        } elseif ($matchEnd !== 0 && (abs($instance) - 1 > ceil(count($split) / 2))) {
            return ExcelError::NA();
        }

        return $split;
    }
}
