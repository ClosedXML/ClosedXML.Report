using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ClosedXML.Report.Utils
{
    [ComVisible(true)]
    [Serializable]
    public class VernoStringReader : TextReader
    {
        private string _s;
        private readonly CultureInfo _culture;
        private int _pos;
        private int _length;

        /// <summary>Инициализирует новый экземпляр класса <see cref="T:System.IO.StringReader" />, осуществляющий чтение из указанной строки.</summary>
        /// <param name="s">Строка, для которой должен быть инициализирован класс <see cref="T:System.IO.StringReader" />. </param>
        /// <exception cref="T:System.ArgumentNullException">Значение параметра <paramref name="s" /> — null. </exception>
        public VernoStringReader(string s): this(s, CultureInfo.CurrentCulture)
        {
        }

        /// <summary>Инициализирует новый экземпляр класса <see cref="T:System.IO.StringReader" />, осуществляющий чтение из указанной строки.</summary>
        /// <param name="s">Строка, для которой должен быть инициализирован класс <see cref="T:System.IO.StringReader" />. </param>
        /// <param name="culture">Culture</param>
        /// <exception cref="T:System.ArgumentNullException">Значение параметра <paramref name="s" /> — null. </exception>
        public VernoStringReader(string s, CultureInfo culture)
        {
            _s = s ?? throw new ArgumentNullException("s");
            _length = s.Length;
            _culture = culture;
        }

        /// <summary>Закрывает объект <see cref="T:System.IO.StringReader" />.</summary>
        /// <filterpriority>2</filterpriority>
        public override void Close()
        {
            Dispose(true);
        }

        /// <summary>Освобождает неуправляемые ресурсы, используемые <see cref="T:System.IO.StringReader" /> (при необходимости освобождает и управляемые ресурсы).</summary>
        /// <param name="disposing">Значение true позволяет освободить управляемые и неуправляемые ресурсы; значение false позволяет освободить только неуправляемые ресурсы. </param>
        protected override void Dispose(bool disposing)
        {
            _s = null;
            _pos = 0;
            _length = 0;
            base.Dispose(disposing);
        }

        /// <summary>Возвращает следующий доступный символ, но не использует его.</summary>
        /// <returns>Целое число, представляющее следующий символ, чтение которого необходимо выполнить, или значение -1, если доступных символов больше нет или поток не поддерживает поиск.</returns>
        /// <exception cref="T:System.ObjectDisposedException">Текущее средство чтения закрыто. </exception>
        /// <filterpriority>2</filterpriority>
        public override int Peek()
        {
            if (_s == null)
                ReaderClosed();
            if (_pos == _length)
                return -1;
            return _s[_pos];
        }

        /// <summary>Считывает следующий символ из строки ввода и увеличивает позицию символа на один символ.</summary>
        /// <returns>Следующий символ из основной строки или значение -1, если больше нет доступных символов.</returns>
        /// <exception cref="T:System.ObjectDisposedException">Текущее средство чтения закрыто. </exception>
        /// <filterpriority>2</filterpriority>
        public override int Read()
        {
            if (_s == null)
                ReaderClosed();
            if (_pos == _length)
                return -1;
            string s = _s;
            int pos = _pos;
            _pos = pos + 1;
            int index = pos;
            return s[index];
        }

        /// <summary>Считывает блок символов из строки ввода и увеличивает позицию символов на <paramref name="count" />.</summary>
        /// <returns>Общее количество символов, считанных в буфер.Оно может быть меньше, чем число запрошенных символов, если большинство символов не доступно в текущий момент, или равно нулю, если достигнут конец основной строки.</returns>
        /// <param name="buffer">При возвращении данного метода содержит заданный массив символов, в котором значения в интервале между <paramref name="index" /> и (<paramref name="index" /> + <paramref name="count" /> - 1) заменены символами, считанными из текущего источника. </param>
        /// <param name="index">Начальный индекс в буфере. </param>
        /// <param name="count">Количество символов, которые необходимо считать. </param>
        /// <exception cref="T:System.ArgumentNullException">Параметр <paramref name="buffer" /> имеет значение null. </exception>
        /// <exception cref="T:System.ArgumentException">Длина буфера за вычетом значения параметра <paramref name="index" /> меньше значения параметра <paramref name="count" />. </exception>
        /// <exception cref="T:System.ArgumentOutOfRangeException">Значение параметра <paramref name="index" /> или <paramref name="count" /> является отрицательным. </exception>
        /// <exception cref="T:System.ObjectDisposedException">Текущее средство чтения закрыто. </exception>
        /// <filterpriority>2</filterpriority>
        public override int Read([In, Out] char[] buffer, int index, int count)
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");
            if (index < 0)
                throw new ArgumentOutOfRangeException("index");
            if (count < 0)
                throw new ArgumentOutOfRangeException("count");
            if (buffer.Length - index < count)
                throw new ArgumentException("Argument invalid off len");
            if (_s == null)
                ReaderClosed();
            int count1 = _length - _pos;
            if (count1 > 0)
            {
                if (count1 > count)
                    count1 = count;
                _s.CopyTo(_pos, buffer, index, count1);
                _pos = _pos + count1;
            }
            return count1;
        }

        /// <summary>Выполняет чтение всех символов, начиная с текущей позиции до конца строки, и возвращает их в виде одной строки.</summary>
        /// <returns>Содержимое, начиная от текущей позиции до конца основной строки.</returns>
        /// <exception cref="T:System.OutOfMemoryException">Недостаточно памяти для размещения буфера возвращаемых строк. </exception>
        /// <exception cref="T:System.ObjectDisposedException">Текущее средство чтения закрыто. </exception>
        /// <filterpriority>2</filterpriority>
        public new VernoStringReader ReadToEnd()
        {
            return new VernoStringReader(base.ReadToEnd(), _culture);
        }

        private string DoReadToEnd()
        {
            if (_s == null)
                ReaderClosed();
            string str = _pos != 0 ? _s.Substring(_pos, _length - _pos) : _s;
            _pos = _length;
            return str;
        }

        /// <summary>Выполняет чтение строки символов из текущей строки и возвращает данные в виде строки.</summary>
        /// <returns>Следующая строка из текущей строки, или значение null, если достигнут конец строки.</returns>
        /// <exception cref="T:System.ObjectDisposedException">Текущее средство чтения закрыто. </exception>
        /// <exception cref="T:System.OutOfMemoryException">Недостаточно памяти для размещения буфера возвращаемых строк. </exception>
        /// <filterpriority>2</filterpriority>
        public new VernoStringReader ReadLine()
        {
            if (_s == null)
                ReaderClosed();
            int pos;
            for (pos = _pos; pos < _length; ++pos)
            {
                char ch = _s[pos];
                switch (ch)
                {
                    case '\r':
                    case '\n':
                        string str = _s.Substring(_pos, pos - _pos);
                        _pos = pos + 1;
                        if (ch != 13 || _pos >= _length || _s[_pos] != 10)
                            return new VernoStringReader(str, _culture);
                        _pos = _pos + 1;
                        return new VernoStringReader(str, _culture);
                    default:
                        continue;
                }
            }
            if (pos <= _pos)
                return null;
            string str1 = _s.Substring(_pos, pos - _pos);
            _pos = pos;
            return new VernoStringReader(str1, _culture);
        }

        public string ReadWord()
        {
            return ReadWhile(char.IsLetterOrDigit);
        }

        public VernoStringReader ReadRegex(Regex regex)
        {
            if (_s == null)
                ReaderClosed();

            var match = regex.Match(_s, _pos);
            if (match.Success)
            {
                _pos = match.Index + match.Value.Length;
                return new VernoStringReader(match.Groups.Count > 1 ? match.Groups[1].Value : match.Value, _culture);
            }
            return new VernoStringReader("", _culture);
        }

        /// <summary>.</summary>
        /// <param name="pattern">Шаблон регулярного выражения для сопоставления. </param>
        public VernoStringReader ReadRegex(string pattern)
        {
            return ReadRegex(new Regex(pattern));
        }

        public VernoStringReader ReadTo(params char[] triggers)
        {
            return ReadTo(triggers.Select(ch => ch.ToString()).ToArray(), false);
        }

        public VernoStringReader ReadTo(params string[] triggers)
        {
            return ReadTo(triggers, false);
        }

        public VernoStringReader ReadTo(string[] triggers, bool closeQuotes)
        {
            if (_s == null)
                ReaderClosed();

            bool quot1 = false;
            bool quot2 = false;
            int start = _pos;
            var slen = _s.Length;
            for (int i = _pos; i < slen; i++)
            {
                if (_s[i] == '\'') quot1 = !quot1;
                if (_s[i] == '"') quot2 = !quot2;

                if (closeQuotes && (quot1 || quot2)) continue;

                if (triggers.Any(pred => slen >= i + pred.Length && _s.Substring(i, pred.Length) == pred))
                {
                    _pos = i + 1;
                    return new VernoStringReader(_s.Substring(start, i - start).Trim(), _culture);
                }
            }
            _pos = slen;
            return new VernoStringReader(_s.Substring(start).Trim(), _culture);
        }

        public VernoStringReader ReadWhile(Predicate<char> whilePredicate)
        {
            return new VernoStringReader(ReadWhile((int idx) => whilePredicate(_s[idx])), _culture);
        }

        private string ReadWhile(Predicate<int> whilePredicate)
        {
            if (_s == null)
                ReaderClosed();

            int start = Find(whilePredicate);
            var slen = _s.Length;
            for (int i = start; i < slen; i++)
            {
                if (!whilePredicate(i))
                {
                    _pos = i;
                    return _s.Substring(start, i - start).Trim();
                }
            }
            _pos = slen;
            return _s.Substring(start).Trim();
        }

        private int Find(Predicate<int> toPredicate)
        {
            var slen = _s.Length;
            for (int i = _pos; i < slen; i++)
            {
                if (toPredicate(i))
                    return i;
            }
            return slen;
        }

        private static readonly Regex IntRegex = new Regex(@"^[+-]?\d+", RegexOptions.IgnoreCase);
        public int ReadInt(Func<string, int> failure = null)
        {
            if (_s == null)
                ReaderClosed();

            SkipSpace();
            var idx = Find(i =>
            {
                var ch = _s[i];
                return char.IsWhiteSpace(ch) || char.IsSeparator(ch) || (ch != '-' && char.IsPunctuation(ch));
            });
            string str = _s.Substring(_pos, idx - _pos);
            var match = IntRegex.Match(str);
            if (match.Success && match.Index <= 1)
            {
                _pos = _pos + match.Index + match.Value.Length;
                return int.Parse(match.Value, _culture);
            }
            else
            {
                if (idx != _s.Length)
                    _pos = idx;
                return failure != null ? failure(str) : 0;
            }
        }

        private static readonly Regex DoubleRegex = new Regex(@"[+-]?\d+([.,]\d+)?", RegexOptions.IgnoreCase);
        public double ReadDouble(Func<string, double> failure = null)
        {
            if (_s == null)
                ReaderClosed();

            SkipSpace();
            var idx = Find(i =>
            {
                var ch = _s[i];
                return char.IsWhiteSpace(ch) || char.IsSeparator(ch) || (ch != '-' && ch != '.' && ch != ',' && char.IsPunctuation(ch));
            });
            string str = _s.Substring(_pos, idx - _pos);
            var match = DoubleRegex.Match(str);
            if (match.Success && match.Index <= 1)
            {
                _pos = _pos + match.Index + match.Value.Length;
                var sep = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
                return double.Parse(match.Value.Replace(".", sep).Replace(",", sep), _culture);
            }
            else
            {
                if (idx != _s.Length)
                    _pos = idx;
                return failure != null ? failure(str) : 0;
            }
        }

        public DateTime ReadDateTime(string dateFormat, Func<string, DateTime> failure = null)
        {
            string str = _s.Substring(_pos, dateFormat.Length);
            _pos += str.Length;

            var dateTimeStyles = DateTimeStyles.AllowWhiteSpaces
                               | DateTimeStyles.AssumeLocal
                               | DateTimeStyles.NoCurrentDateDefault;
            return DateTime.TryParseExact(str, dateFormat, _culture, dateTimeStyles, out var result)
                ? result
                : (failure?.Invoke(str) ?? DateTime.MinValue);
        }

        private static readonly Regex DateTimeRegex = new Regex(@"([1-9]|0[1-9]|[12][0-9]|3[01])\D([1-9]|0[1-9]|1[012])\D(19[0-9][0-9]|20[0-9][0-9])", RegexOptions.IgnoreCase);

        public DateTime ReadDateTime(Func<string, DateTime> failure = null)
        {
            if (_s == null)
                ReaderClosed();

            string str = _s.Substring(_pos, _s.Length - _pos);
            var match = DateTimeRegex.Match(str);
            if (match.Success && match.Index <= 1)
            {
                var dateTimeStyles = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal | DateTimeStyles.NoCurrentDateDefault;
                _pos = _pos + match.Index + match.Value.Length;
                return DateTime.Parse(match.Value, _culture, dateTimeStyles);
            }
            else
                return failure != null ? failure(str) : DateTime.MinValue;
        }

#if !NET40
        /// <summary>Асинхронно выполняет чтение строки символов из текущей строки и возвращает данные в виде строки.</summary>
        /// <returns>Задача, представляющая асинхронную операцию чтения.Значение параметра <paramref name="TResult" /> содержит следующую строку из средства чтения строк или значение null, если все знаки считаны.</returns>
        /// <exception cref="T:System.ArgumentOutOfRangeException">Количество символов в следующей строке больше <see cref="F:System.Int32.MaxValue" />.</exception>
        /// <exception cref="T:System.ObjectDisposedException">Удалено средство чтения строки.</exception>
        /// <exception cref="T:System.InvalidOperationException">Средство чтения в настоящее время используется предыдущей операцией чтения. </exception>
        [ComVisible(false)]
        public override Task<string> ReadLineAsync()
        {
            return Task.FromResult(ReadLine().ToString());
        }

        /// <summary>Асинхронно считывает все символы, начиная с текущей позиции до конца строки, и возвращает их в виде одной строки.</summary>
        /// <returns>Задача, представляющая асинхронную операцию чтения.Значение параметра <paramref name="TResult" /> содержит строку с символами от текущего положения до конца строки.</returns>
        /// <exception cref="T:System.ArgumentOutOfRangeException">Количество символов больше <see cref="F:System.Int32.MaxValue" />.</exception>
        /// <exception cref="T:System.ObjectDisposedException">Удалено средство чтения строки.</exception>
        /// <exception cref="T:System.InvalidOperationException">Средство чтения в настоящее время используется предыдущей операцией чтения. </exception>
        [ComVisible(false)]
        public override Task<string> ReadToEndAsync()
        {
            return Task.FromResult(ReadToEnd().ToString());
        }

        /// <summary>Асинхронно считывает указанное максимальное количество символов из текущей строки и записывает данные в буфер, начиная с заданного индекса.</summary>
        /// <returns>Задача, представляющая асинхронную операцию чтения.Значение параметра <paramref name="TResult" /> содержит общее число байтов, считанных в буфер.Значение результата может быть меньше запрошенного числа байтов, если число текущих доступных байтов меньше запрошенного числа, или результат может быть равен 0 (нулю), если был достигнут конец строки.</returns>
        /// <param name="buffer">При возвращении данного метода содержит заданный массив символов, в котором значения в интервале между <paramref name="index" /> и (<paramref name="index" /> + <paramref name="count" /> - 1) заменены символами, считанными из текущего источника.</param>
        /// <param name="index">Позиция в буфере <paramref name="buffer" />, с которого начинается запись.</param>
        /// <param name="count">Наибольшее число символов для чтения.Если конец строки достигнут, прежде чем в буфер записано указанное количество символов, метод возвращает управление.</param>
        /// <exception cref="T:System.ArgumentNullException">Параметр <paramref name="buffer" /> имеет значение null.</exception>
        /// <exception cref="T:System.ArgumentOutOfRangeException">Значение параметра <paramref name="index" /> или <paramref name="count" /> является отрицательным.</exception>
        /// <exception cref="T:System.ArgumentException">Сумма значений параметров <paramref name="index" /> и <paramref name="count" /> больше длины буфера.</exception>
        /// <exception cref="T:System.ObjectDisposedException">Удалено средство чтения строки.</exception>
        /// <exception cref="T:System.InvalidOperationException">Средство чтения в настоящее время используется предыдущей операцией чтения. </exception>
        [ComVisible(false)]
        public override Task<int> ReadBlockAsync(char[] buffer, int index, int count)
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");
            if (index < 0 || count < 0)
                throw new ArgumentOutOfRangeException(index < 0 ? "index" : "count");
            if (buffer.Length - index < count)
                throw new ArgumentException("Argument invalid off len");
            return Task.FromResult(ReadBlock(buffer, index, count));
        }

        /// <summary>Асинхронно считывает указанное максимальное количество символов из текущей строки и записывает данные в буфер, начиная с заданного индекса. </summary>
        /// <returns>Задача, представляющая асинхронную операцию чтения.Значение параметра <paramref name="TResult" /> содержит общее число байтов, считанных в буфер.Значение результата может быть меньше запрошенного числа байтов, если число текущих доступных байтов меньше запрошенного числа, или результат может быть равен 0 (нулю), если был достигнут конец строки.</returns>
        /// <param name="buffer">При возвращении данного метода содержит заданный массив символов, в котором значения в интервале между <paramref name="index" /> и (<paramref name="index" /> + <paramref name="count" /> - 1) заменены символами, считанными из текущего источника.</param>
        /// <param name="index">Позиция в буфере <paramref name="buffer" />, с которого начинается запись.</param>
        /// <param name="count">Наибольшее число символов для чтения.Если конец строки достигнут, прежде чем в буфер записано указанное количество символов, метод возвращает управление.</param>
        /// <exception cref="T:System.ArgumentNullException">Параметр <paramref name="buffer" /> имеет значение null.</exception>
        /// <exception cref="T:System.ArgumentOutOfRangeException">Значение параметра <paramref name="index" /> или <paramref name="count" /> является отрицательным.</exception>
        /// <exception cref="T:System.ArgumentException">Сумма значений параметров <paramref name="index" /> и <paramref name="count" /> больше длины буфера.</exception>
        /// <exception cref="T:System.ObjectDisposedException">Удалено средство чтения строки.</exception>
        /// <exception cref="T:System.InvalidOperationException">Средство чтения в настоящее время используется предыдущей операцией чтения. </exception>
        [ComVisible(false)]
        public override Task<int> ReadAsync(char[] buffer, int index, int count)
        {
            if (buffer == null)
                throw new ArgumentNullException("buffer");
            if (index < 0 || count < 0)
                throw new ArgumentOutOfRangeException(index < 0 ? "index" : "count");
            if (buffer.Length - index < count)
                throw new ArgumentException("Argument invalid off len");
            return Task.FromResult(Read(buffer, index, count));
        }
#endif

        internal static void ReaderClosed()
        {
            throw new ObjectDisposedException(null, "Object disposed. Reader is closed.");
        }

        public void SkipSpace()
        {
            if (_s == null)
                ReaderClosed();

            var slen = _s.Length;
            while (_pos < slen && char.IsWhiteSpace(_s[_pos]))
            {
                _pos++;
            }
        }

        public void Skip(params char[] chars)
        {
            if (_s == null)
                ReaderClosed();

            var slen = _s.Length;
            while (_pos<slen && Array.IndexOf(chars, _s[_pos]) >= 0)
            {
                _pos++;
            }
        }

        public VernoStringReader ReadInBrackets(char openBracket, char closeBracket = Char.MinValue)
        {
            if (_s == null)
                ReaderClosed();
            if (closeBracket == char.MinValue)
            {
                closeBracket = openBracket;
            }
            bool sameBr = closeBracket == openBracket;

            ReadTo(openBracket);
            int cnt = 1;
            int start = _pos;
            do
            {
                ReadTo(openBracket, closeBracket);
                if (LastChar == openBracket)
                    cnt++;
                else
                    cnt--;
            } while (!sameBr && cnt>0);
            if (LastChar != closeBracket)
                _pos++;
            return new VernoStringReader(_s.Substring(start, _pos-start-1).Trim(), _culture);
        }

        public char LastChar
        {
            get { return _s[_pos-1]; }
        }

        public override string ToString()
        {
            if (_s == null)
                ReaderClosed();
            return _s;
        }

        public static implicit operator string(VernoStringReader reader)
        {
            return reader.DoReadToEnd();
        }

        public static implicit operator VernoStringReader(string str)
        {
            return new VernoStringReader(str, CultureInfo.CurrentCulture);
        }

        public char ReadChar()
        {
            return (char) Read();
        }

        public string[] ReadArray(params char[] separator)
        {
            return ReadArray(separator.Select(ch=>ch.ToString()).ToArray());
        }

        public string[] ReadArray(params string[] separator)
        {
            return Split(separator).Select(s => s.ToString().Trim('"', '\'')).ToArray();
        }

        public T[] ReadArray<T>(params char[] separator)
        {
            return ReadArray<T>(separator.Select(ch=>ch.ToString()).ToArray());
        }

        public T[] ReadArray<T>(params string[] separator)
        {
            var strArr = ReadArray(separator);
            return strArr.Select(str => (T)str.ChangeType(typeof(T), _culture)).ToArray();
        }

        public VernoStringReader[] Split(params string[] separator)
        {
            if (_s == null)
                ReaderClosed();

            Skip(separator.Select(c=>c[0]).ToArray());
            var result = new List<VernoStringReader>();
            var slen = _s.Length;
            while (_pos < slen)
            {
                result.Add(ReadTo(separator, true));
            }
            return result.ToArray();
        }

        public KeyValuePair<string, string>[] ReadNamedValues(string pairSeparator, string assignChar)
        {
            var arr = Split(pairSeparator);
            return (from pair in arr
                    let key = pair.ReadTo(new[] {assignChar}, true).ToString().Trim('"', '\'')
                    let value = pair.DoReadToEnd().Trim('"', '\'')
                    select new KeyValuePair<string, string>(key, value))
                .ToArray();
        }
    }
}
