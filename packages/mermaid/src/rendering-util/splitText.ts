import type { CheckFitFunction, MarkdownLine, MarkdownWord, MarkdownWordType } from './types.js';

/**
 * Splits a string into graphemes if available, otherwise characters.
 */
export function splitTextToChars(text: string): string[] {
  if (Intl.Segmenter) {
    return [...new Intl.Segmenter().segment(text)].map((s) => s.segment);
  }
  return [...text];
}

/**
 * Splits a string into words by using `Intl.Segmenter` if available, or splitting by ' '.
 * `Intl.Segmenter` uses the default locale, which might be different across browsers.
 */
export function splitLineToWords(text: string): string[] {
  if (Intl.Segmenter) {
    return [...new Intl.Segmenter(undefined, { granularity: 'word' }).segment(text)].map(
      (s) => s.segment
    );
  }
  // Split by ' ' removes the ' 's from the result.
  const words = text.split(' ');
  // Add the ' 's back to the result.
  const wordsWithSpaces = words.flatMap((s) => [s, ' ']).filter((s) => s);
  // Remove last space.
  wordsWithSpaces.pop();
  return wordsWithSpaces;
}

/**
 * Splits a word into two parts, the first part fits the width and the remaining part.
 * @param checkFit - Function to check if word fits
 * @param word - Word to split
 * @returns [first part of word that fits, rest of word]
 */
export function splitWordToFitWidth(
  checkFit: CheckFitFunction,
  word: MarkdownWord
): [MarkdownWord, MarkdownWord] {
  const characters = splitTextToChars(word.content);
  return splitWordToFitWidthRecursion(checkFit, [], characters, word.type);
}

function splitWordToFitWidthRecursion(
  checkFit: CheckFitFunction,
  usedChars: string[],
  remainingChars: string[],
  type: MarkdownWordType
): [MarkdownWord, MarkdownWord] {
  const allChars = [...usedChars, ...remainingChars];
  if (allChars.length === 0) {
    return [
      { content: '', type },
      { content: '', type },
    ];
  }

  // Use binary search to find how many characters fit,
  // reducing O(n) checkFit calls to O(log n)
  const start = usedChars.length;
  let lo = start;
  let hi = allChars.length;
  let bestFit = start; // number of chars that fit

  // Check if all characters fit
  if (checkFit([{ content: allChars.join(''), type }])) {
    return [
      { content: allChars.join(''), type },
      { content: '', type },
    ];
  }

  while (lo <= hi) {
    const mid = Math.floor((lo + hi) / 2);
    if (mid === 0) {
      lo = 1;
      continue;
    }
    const candidate = allChars.slice(0, mid).join('');
    if (checkFit([{ content: candidate, type }])) {
      bestFit = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }

  // If no character fits, force at least one character to avoid infinite loops
  if (bestFit === 0) {
    bestFit = 1;
  }

  return [
    { content: allChars.slice(0, bestFit).join(''), type },
    { content: allChars.slice(bestFit).join(''), type },
  ];
}

/**
 * Splits a line into multiple lines that satisfy the checkFit function.
 * @param line - Line to split
 * @param checkFit - Function to check if line fits
 * @returns Array of lines that fit
 */
export function splitLineToFitWidth(
  line: MarkdownLine,
  checkFit: CheckFitFunction
): MarkdownLine[] {
  if (line.some(({ content }) => content.includes('\n'))) {
    throw new Error('splitLineToFitWidth does not support newlines in the line');
  }
  return splitLineToFitWidthRecursion(line, checkFit);
}

function splitLineToFitWidthRecursion(
  words: MarkdownWord[],
  checkFit: CheckFitFunction,
  lines: MarkdownLine[] = [],
  newLine: MarkdownLine = []
): MarkdownLine[] {
  // Return if there is nothing left to split
  if (words.length === 0) {
    // If there is a new line, add it to the lines
    if (newLine.length > 0) {
      lines.push(newLine);
    }
    return lines.length > 0 ? lines : [];
  }
  let joiner = '';
  if (words[0].content === ' ') {
    joiner = ' ';
    words.shift();
  }
  const nextWord: MarkdownWord = words.shift() ?? { content: ' ', type: 'normal' };
  const lineWithNextWord: MarkdownLine = [...newLine];
  if (joiner !== '') {
    lineWithNextWord.push({ content: joiner, type: 'normal' });
  }
  lineWithNextWord.push(nextWord);

  if (checkFit(lineWithNextWord)) {
    // nextWord fits, so we can add it to the new line and continue
    return splitLineToFitWidthRecursion(words, checkFit, lines, lineWithNextWord);
  }

  // nextWord doesn't fit, so we need to split it
  if (newLine.length > 0) {
    // There was text in newLine, so add it to lines and push nextWord back into words.
    lines.push(newLine);
    words.unshift(nextWord);
  } else if (nextWord.content) {
    // There was no text in newLine, so we need to split nextWord
    const [line, rest] = splitWordToFitWidth(checkFit, nextWord);
    lines.push([line]);
    if (rest.content) {
      words.unshift(rest);
    }
  }
  return splitLineToFitWidthRecursion(words, checkFit, lines);
}
