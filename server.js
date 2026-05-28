const express = require('express');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const { GoogleGenAI, GoogleGenAIError } = require('@google/generative-ai');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Path definitions
const EXCEL_PATH = path.join(__dirname, 'Topics.xlsx');
const DB_PATH = path.join(__dirname, 'journey_db.json');

// Predefined learning bank (rich expert content) to guarantee the app has a premium, rich learning system immediately
const DEFAULT_LEARNING_BANK = {
  "AI": [
    {
      "id": "ai_1",
      "question": "What is Artificial Intelligence (AI) and its primary subfields?",
      "answer": "Artificial Intelligence is the branch of computer science focused on building smart systems capable of performing tasks that typically require human intelligence, such as visual perception, decision-making, and translation. Its main subfields include Machine Learning (ML), Deep Learning (DL), Natural Language Processing (NLP), Computer Vision, and Robotics."
    },
    {
      "id": "ai_2",
      "question": "What is the difference between Narrow AI and General AI (AGI)?",
      "answer": "Narrow AI (Weak AI) is trained and focused to perform a single specific task, such as translation or playing chess (e.g., Siri, AlphaGo). General AI (Strong AI / AGI) is a hypothetical system that possesses human-like cognitive abilities across all domains, enabling it to learn, reason, and apply intelligence to any problem it encounters."
    }
  ],
  "ML": [
    {
      "id": "ml_1",
      "question": "Explain Supervised, Unsupervised, and Reinforcement Learning.",
      "answer": "1. **Supervised Learning**: Models are trained on labeled data (input-output pairs), mapping inputs to known correct targets (e.g., predicting house prices, email spam classification).\n2. **Unsupervised Learning**: Models identify hidden patterns and groupings in unlabeled data without explicit target answers (e.g., customer segmentation via K-Means clustering).\n3. **Reinforcement Learning**: An agent learns to make decisions by performing actions in an environment to maximize cumulative rewards through trial and error (e.g., training self-driving cars or game-playing agents)."
    },
    {
      "id": "ml_2",
      "question": "What is Overfitting and how can it be prevented?",
      "answer": "Overfitting occurs when a machine learning model learns the training data's noise and details too well, causing it to perform exceptionally on training sets but poorly on unseen test data. It can be prevented by:\n- Using **regularization** techniques (L1/L2 penalties).\n- Gathering more training data.\n- Applying **cross-validation**.\n- Using **dropout layers** (in neural nets) or pruning (in decision trees).\n- Simplifying the model architecture."
    }
  ],
  "NLP": [
    {
      "id": "nlp_1",
      "question": "What is Tokenization and why is it crucial in NLP?",
      "answer": "Tokenization is the process of breaking down a stream of text (sentences or documents) into smaller, manageable units called **tokens** (which can be words, characters, or subwords like Byte-Pair Encoding tokens). It is the critical first step in NLP because algorithms require structured numerical inputs, and tokenization allows raw text to be mapped to vocabulary indices."
    },
    {
      "id": "nlp_2",
      "question": "What are Word Embeddings and how do they capture meaning?",
      "answer": "Word Embeddings (e.g., Word2Vec, GloVe, BERT embeddings) are dense vector representations of words in a high-dimensional continuous space. They capture semantic meanings by mapping words that appear in similar contexts close together in the vector space. This allows mathematical capture of word relationships, like 'king - man + woman = queen'."
    }
  ],
  "DSA": [
    {
      "id": "dsa_found_1",
      "question": "1. Intro to DSA & Complexity Theory",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Data Structures organize and store data efficiently, while Algorithms provide step-by-step instructions to process it. Big O Notation measures the worst-case scaling rate of Execution Time and Space (memory) relative to the input size $N$.</p><strong>Scale of Runtimes (Best to Worst):</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li><strong>O(1) - Constant</strong>: Math operations, direct array lookup.</li><li><strong>O(log N) - Logarithmic</strong>: Binary search (halving search space).</li><li><strong>O(N) - Linear</strong>: Single pass loop over a collection.</li><li><strong>O(N log N) - Linearithmic</strong>: Optimal sorting algorithms (Merge, Quick Sort average).</li><li><strong>O(N^2) - Quadratic</strong>: Nested loops, e.g., Bubble Sort, Matrix scans.</li><li><strong>O(2^N) - Exponential</strong>: Double-branching recursions, e.g., Fibonacci.</li></ul>"
    },
    {
      "id": "dsa_found_2",
      "question": "2. Mathematics (GCD, Primes, Sieve & Bitwise)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Primes form the bedrock of optimal distributions and cryptography. Prime checks run in $O(\\sqrt{N})$; Sieve of Eratosthenes generates primes up to $N$ in $O(N \\log \\log N)$ time. Bit Manipulation runs directly on registers using binary operations (AND, OR, XOR, NOT, Shifts) in a single CPU cycle.</p><strong>Code Implementation (GCD, Bitwise Power of 2, Sieve):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">math_utilities.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">// Euclidean Greatest Common Divisor\nfunction gcd(a, b) {\n  return b === 0 ? a : gcd(b, a % b);\n}\n\n// Fast Bitwise Power-of-Two Validation\nfunction isPowerOfTwo(n) {\n  return n > 0 && (n & (n - 1)) === 0;\n}\n\n// Sieve of Eratosthenes Prime Generator\nfunction sieve(n) {\n  let isPrime = Array(n + 1).fill(true);\n  isPrime[0] = isPrime[1] = false;\n  for (let p = 2; p * p <= n; p++) {\n    if (isPrime[p]) {\n      for (let i = p * p; i <= n; i += p) isPrime[i] = false;\n    }\n  }\n  let primes = [];\n  for (let i = 2; i <= n; i++) {\n    if (isPrime[i]) primes.push(i);\n  }\n  return primes;\n}</pre></div>"
    },
    {
      "id": "dsa_found_3",
      "question": "3. Recursion & Backtracking (N-Queens Solver)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Recursion calls functions self-similarly down to base cases. Backtracking is an exhaustive DFS traversal that tests options step-by-step, pruning invalid branches, and rolling back (backtracking) changes if a violation occurs.</p><strong>Code Implementation (N-Queens Backtracking Solver):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">n_queens.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function solveNQueens(n) {\n  let cols = new Set(), diag1 = new Set(), diag2 = new Set();\n  let board = Array(n).fill().map(() => Array(n).fill(\".\"));\n  let results = [];\n  \n  function backtrack(row) {\n    if (row === n) {\n      results.push(board.map(r => r.join(\"\")));\n      return;\n    }\n    for (let col = 0; col < n; col++) {\n      let d1 = row - col;\n      let d2 = row + col;\n      if (cols.has(col) || diag1.has(d1) || diag2.has(d2)) continue;\n      \n      board[row][col] = \"Q\";\n      cols.add(col); diag1.add(d1); diag2.add(d2);\n      \n      backtrack(row + 1);\n      \n      board[row][col] = \".\"; // Backtrack / revert choice\n      cols.delete(col); diag1.delete(d1); diag2.delete(d2);\n    }\n  }\n  \n  backtrack(0);\n  return results;\n}</pre></div>"
    },
    {
      "id": "dsa_found_4",
      "question": "4. Greedy Algorithms (Job Scheduling & Change)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Greedy algorithms make locally optimal decisions at each branch hoping for a global optimum. They resolve optimization challenges if problems satisfy the Greedy Choice Property and Optimal Substructure (e.g. Huffman coding, Job Scheduling, Coin Change).</p><strong>Code Implementation (Job Scheduling with Deadlines to Maximize Profit):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">job_scheduling.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function scheduleJobs(jobs) {\n  // jobs: [{ id, deadline, profit }]\n  // Sort jobs descending by profit\n  jobs.sort((a, b) => b.profit - a.profit);\n  \n  let maxDeadline = Math.max(...jobs.map(j => j.deadline));\n  let slots = Array(maxDeadline + 1).fill(null);\n  let totalProfit = 0;\n  \n  for (let job of jobs) {\n    // Greedy choice: schedule in latest free slot before deadline\n    for (let t = job.deadline; t > 0; t--) {\n      if (slots[t] === null) {\n        slots[t] = job.id;\n        totalProfit += job.profit;\n        break;\n      }\n    }\n  }\n  return { slots: slots.filter(s => s !== null), totalProfit };\n}</pre></div>"
    },
    {
      "id": "dsa_found_5",
      "question": "5. Strings (Anagrams, Palindromes & Z-Algorithm)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Strings represent character sequences. Matching algorithms analyze string matches: KMP precomputes prefix values; Z-Algorithm builds a match-count array Z where Z[i] is the length of the longest common prefix between suffix starting at i and prefix of the string.</p><strong>Code Implementation (Z-Algorithm Pattern Finder):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">z_algorithm.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function getZArray(str) {\n  let n = str.length, z = Array(n).fill(0);\n  let l = 0, r = 0;\n  for (let i = 1; i < n; i++) {\n    if (i <= r) {\n      z[i] = Math.min(r - i + 1, z[i - l]);\n    }\n    while (i + z[i] < n && str[z[i]] === str[i + z[i]]) {\n      z[i]++;\n    }\n    if (i + z[i] - 1 > r) {\n      l = i; r = i + z[i] - 1;\n    }\n  }\n  return z;\n}\n\nfunction searchZ(text, pattern) {\n  let concat = pattern + \"$\" + text;\n  let z = getZArray(concat);\n  let results = [];\n  for (let i = 0; i < z.length; i++) {\n    if (z[i] === pattern.length) {\n      results.push(i - pattern.length - 1); // Match index in text\n    }\n  }\n  return results;\n}</pre></div>"
    },
    {
      "id": "dsa_found_6",
      "question": "6. Stacks & Queues (Infix to Postfix Converter)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">1. **Stack**: Last-In-First-Out (LIFO) structure. Evaluates expressions and handles recursion tracks.\n2. **Queue**: First-In-First-Out (FIFO) structure. Manages scheduling buffers.\nInfix to Postfix parsing (Shunting-yard algorithm) parses human-readable arithmetic into machine-computable instructions.</p><strong>Code Implementation (Infix to Postfix Converter):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">infix_shunting.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function infixToPostfix(exp) {\n  let precedence = { \"+\": 1, \"-\": 1, \"*\": 2, \"/\": 2, \"^\": 3 };\n  let stack = [], result = [];\n  for (let char of exp) {\n    if (/[a-zA-Z0-9]/.test(char)) result.push(char);\n    else if (char === \"(\") stack.push(char);\n    else if (char === \")\") {\n      while (stack.length && stack[stack.length - 1] !== \"(\") {\n        result.push(stack.pop());\n      }\n      stack.pop(); // Pop '('\n    } else {\n      while (stack.length && precedence[char] <= precedence[stack[stack.length - 1]]) {\n        result.push(stack.pop());\n      }\n      stack.push(char);\n    }\n  }\n  while (stack.length) result.push(stack.pop());\n  return result.join(\"\");\n}</pre></div>"
    },
    {
      "id": "dsa_found_7",
      "question": "7. Trees & BST (In/Pre/Post Traversals)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Hierarchical node configurations. A BST satisfies: Left child < Parent < Right child. Traversals order nodes: Inorder (Left, Parent, Right - returns sorted values for BST), Preorder (Parent, Left, Right - used for copying), and Postorder (Left, Right, Parent - used for deletions).</p><strong>Code Implementation (Binary Tree Traversals - Iterative or Recursive DFS):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">tree_traversals.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">class TreeNode {\n  constructor(val) { this.val = val; this.left = this.right = null; }\n}\n\nfunction getTraversals(root) {\n  let inorder = [], preorder = [], postorder = [];\n  \n  function dfs(node) {\n    if (!node) return;\n    preorder.push(node.val); // Preorder: Parent-Left-Right\n    dfs(node.left);\n    inorder.push(node.val);  // Inorder: Left-Parent-Right\n    dfs(node.right);\n    postorder.push(node.val); // Postorder: Left-Right-Parent\n  }\n  \n  dfs(root);\n  return { preorder, inorder, postorder };\n}</pre></div>"
    },
    {
      "id": "dsa_found_8",
      "question": "8. Searching (Binary Search on Answer)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Binary Search on Answer solves optimization challenges where the result space is monotonic. Instead of lookup indices, it binary-searches the threshold value range and validates feasibility at each step (e.g. Koko Eating Bananas, Allocating Books).</p><strong>Code Implementation (Koko Eating Bananas - O(N log max_pile) runtime):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">koko_bananas.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function minEatingSpeed(piles, h) {\n  let left = 1, right = Math.max(...piles);\n  let optimalSpeed = right;\n  \n  function canEatAll(speed) {\n    let hours = 0;\n    for (let p of piles) hours += Math.ceil(p / speed);\n    return hours <= h;\n  }\n  \n  while (left <= right) {\n    let speed = Math.floor((left + right) / 2);\n    if (canEatAll(speed)) {\n      optimalSpeed = speed; // Feasible speed, seek smaller speeds\n      right = speed - 1;\n    } else {\n      left = speed + 1; // Needs faster eating rate\n    }\n  }\n  return optimalSpeed;\n}</pre></div>"
    },
    {
      "id": "dsa_found_9",
      "question": "9. Tries (Insert, Search & Autocomplete)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">A Trie (Prefix Tree) is an advanced retrieval tree node structure where edge mappings represent characters. It validates insertions and lookups in $O(L)$ time (length of string $L$), enabling rapid dictionary lookups and prefix auto-completes.</p><strong>Code Implementation (Trie Auto-complete Engine):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">trie_autocomplete.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">class TrieNode {\n  constructor() { this.children = {}; this.isEnd = false; }\n}\n\nclass AutocompleteTrie {\n  constructor() { this.root = new TrieNode(); }\n  \n  insert(word) {\n    let current = this.root;\n    for (let char of word) {\n      if (!current.children[char]) current.children[char] = new TrieNode();\n      current = current.children[char];\n    }\n    current.isEnd = true;\n  }\n  \n  getSuggestions(prefix) {\n    let current = this.root;\n    for (let char of prefix) {\n      if (!current.children[char]) return [];\n      current = current.children[char];\n    }\n    let results = [];\n    this.dfs(current, prefix, results);\n    return results;\n  }\n  \n  dfs(node, wordAccumulator, results) {\n    if (node.isEnd) results.push(wordAccumulator);\n    for (let char in node.children) {\n      this.dfs(node.children[char], wordAccumulator + char, results);\n    }\n  }\n}</pre></div>"
    },
    {
      "id": "dsa_found_10",
      "question": "10. Linked Lists (Floyd's Cycle Start Detector)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Non-contiguous node chains linked via pointer references. Singly lists link forward; Doubly lists link bidirectionally. Floyd's Tortoise and Hare detects loop cycles and isolates cycle start points using mathematical relative speeds.</p><strong>Code Implementation (Floyd's Cycle Entry Detector):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">linked_list_cycle.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function detectCycleEntry(head) {\n  let slow = head, fast = head;\n  while (fast && fast.next) {\n    slow = slow.next;\n    fast = fast.next.next;\n    if (slow === fast) { // Intersection loop found!\n      // Find loop cycle starting entry node\n      let pointer = head;\n      while (pointer !== slow) {\n        pointer = pointer.next;\n        slow = slow.next;\n      }\n      return pointer; // Entry node starting cycle\n    }\n  }\n  return null; // Cyclic path does not exist\n}</pre></div>"
    },
    {
      "id": "dsa_found_11",
      "question": "11. Dynamic Programming (Longest Increasing Subsequence)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">DP stores recurring sub-problem outcomes to resolve overlapping calculations. Techniques utilize Memoization (Top-down recursion) or Tabulation (Bottom-up grids). Standard DP sequences include LCS and LIS.</p><strong>Code Implementation (LIS - O(N log N) Binary Search DP):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">longest_subsequence.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function lengthOfLIS(nums) {\n  let tails = [];\n  for (let num of nums) {\n    // Binary search to find slot index to insert/replace num\n    let left = 0, right = tails.length - 1, slot = tails.length;\n    while (left <= right) {\n      let mid = Math.floor((left + right) / 2);\n      if (tails[mid] >= num) {\n        slot = mid; right = mid - 1;\n      } else {\n        left = mid + 1;\n      }\n    }\n    if (slot === tails.length) tails.push(num);\n    else tails[slot] = num; // Replace value preserving monotonically\n  }\n  return tails.length;\n}</pre></div>"
    },
    {
      "id": "dsa_found_12",
      "question": "12. Arrays & Sorting (Kadane's & Matrix Rotation)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Contiguous array storage. Prefix sum vectors compute query indexes in $O(1)$ time. Kadane's algorithm calculates maximum subarray sums linearly in $O(N)$ time. 2D coordinates matrix rotations swap index variables in-place.</p><strong>Code Implementation (In-Place 90-degree Matrix Rotation Clockwise):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">matrix_rotation.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function rotateMatrix(matrix) {\n  let n = matrix.length;\n  // 1. Transpose the matrix (swap values across key diagonal)\n  for (let i = 0; i < n; i++) {\n    for (let j = i + 1; j < n; j++) {\n      let temp = matrix[i][j];\n      matrix[i][j] = matrix[j][i];\n      matrix[j][i] = temp;\n    }\n  }\n  // 2. Reverse each row to finalize clockwise rotation\n  for (let i = 0; i < n; i++) {\n    matrix[i].reverse();\n  }\n}</pre></div>"
    },
    {
      "id": "dsa_found_13",
      "question": "13. Hashing (HashMap Linear Probing Prober)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Hashing maps key assets to coordinate index array addresses. When separate keys resolve to identical index targets, collisions occur. Resolving systems use Chaining (list buckets at indices) or Open Addressing (sequential probes like Linear Probing).</p><strong>Code Implementation (Linear Probing HashTable):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">linear_probe.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">class LinearProbingHashTable {\n  constructor(size = 31) {\n    this.keys = Array(size).fill(null);\n    this.vals = Array(size).fill(null);\n  }\n  \n  hash(key) { return Math.abs(key) % this.keys.length; }\n  \n  put(key, val) {\n    let idx = this.hash(key);\n    while (this.keys[idx] !== null) {\n      if (this.keys[idx] === key) { this.vals[idx] = val; return; }\n      idx = (idx + 1) % this.keys.length; // Linear Probe shift\n    }\n    this.keys[idx] = key; this.vals[idx] = val;\n  }\n  \n  get(key) {\n    let idx = this.hash(key);\n    while (this.keys[idx] !== null) {\n      if (this.keys[idx] === key) return this.vals[idx];\n      idx = (idx + 1) % this.keys.length;\n    }\n    return null; // Not found\n  }\n}</pre></div>"
    },
    {
      "id": "dsa_found_14",
      "question": "14. Graphs & MST (Kruskal's & Disjoint Set DSU)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Graphs map nodes connected via edges. A Minimum Spanning Tree (MST) links all nodes using edges of minimum cumulative weights. Kruskal's algorithm greedily collects edges using a Disjoint Set Union (DSU) with Path Compression to prevent cycles in $O(E \\log V)$ time.</p><strong>Code Implementation (DSU & Kruskal's MST):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">kruskals_dsu.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">class DSU {\n  constructor(n) {\n    this.parent = Array(n).fill(0).map((_, i) => i);\n    this.rank = Array(n).fill(0);\n  }\n  find(i) {\n    if (this.parent[i] === i) return i;\n    return this.parent[i] = this.find(this.parent[i]); // Path compression\n  }\n  union(i, j) {\n    let rootI = this.find(i), rootJ = this.find(j);\n    if (rootI !== rootJ) {\n      if (this.rank[rootI] < this.rank[rootJ]) this.parent[rootI] = rootJ;\n      else if (this.rank[rootI] > this.rank[rootJ]) this.parent[rootJ] = rootI;\n      else { this.parent[rootJ] = rootI; this.rank[rootI]++; }\n      return true;\n    }\n    return false;\n  }\n}\n\nfunction kruskalMST(numNodes, edges) {\n  // edges: [{ u, v, weight }]\n  edges.sort((a, b) => a.weight - b.weight);\n  let dsu = new DSU(numNodes), mst = [], totalCost = 0;\n  for (let edge of edges) {\n    if (dsu.union(edge.u, edge.v)) {\n      mst.push(edge);\n      totalCost += edge.weight;\n    }\n  }\n  return { mst, totalCost };\n}</pre></div>"
    },
    {
      "id": "dsa_found_15",
      "question": "15. Core Templates (DFS, Binary Search, DP Knapsack)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Keep this quick reference checklist of standard, robust coding templates commonly expected during software engineering technical interviews.</p><strong>Common Binary Search & DFS Recursion Templates:</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">reference_templates.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">// Binary Search standard interval lookup\nfunction binarySearch(arr, target) {\n  let left = 0, right = arr.length - 1;\n  while (left <= right) {\n    let mid = Math.floor((left + right) / 2);\n    if (arr[mid] === target) return mid;\n    else if (arr[mid] < target) left = mid + 1;\n    else right = mid - 1;\n  }\n  return -1;\n}\n\n// Stateful DFS Recursion Skeleton\nfunction dfsTraversal(node, visitedSet) {\n  if (!node || visitedSet.has(node.id)) return;\n  visitedSet.add(node.id);\n  process(node);\n  for (let neighbor of node.neighbors) {\n    dfsTraversal(neighbor, visitedSet);\n  }\n}</pre></div>"
    },
    {
      "id": "dsa_found_16",
      "question": "16. Mini Projects (Visualizers & Animators)",
      "answer": "<strong>Core Concepts:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Build real applications for portfolio highlights. Creating pathfinding simulators (BFS/Dijkstra matrix animations) or sorting controllers relies on capturing snapshot indices at each algorithm step and rendering them sequentially via intervals.</p><strong>Visualizer Framework Concept:</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">visualizer_core.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">// Core step recorder for visual animations\nclass StepVisualizer {\n  constructor() { this.snapshots = []; }\n  \n  recordState(arrayState, highlightedIndices) {\n    this.snapshots.push({\n      array: [...arrayState],\n      active: [...highlightedIndices]\n    });\n  }\n  \n  animate(renderCallback, intervalDelay = 100) {\n    let stepIndex = 0;\n    const loop = setInterval(() => {\n      if (stepIndex >= this.snapshots.length) return clearInterval(loop);\n      renderCallback(this.snapshots[stepIndex++]);\n    }, intervalDelay);\n  }\n}</pre></div>"
    },
    {
      "id": "dsa_found_17",
      "question": "17. Beginner Coding Mistakes & Validation Guards",
      "answer": "<strong>Common Mistakes:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Write defensive code by adding safety guards and avoiding typical traps:</p><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li><strong>Assignment vs. Comparison</strong>: Writing <code>if (x = y)</code> performs an assignment, returning the value of y. Always use <code>if (x === y)</code>.</li><li><strong>Forgetting Semicolons</strong>: While JS uses ASI, omitting semicolons in languages like Java/C/C++ causes compile failures. Check line endings carefully.</li><li><strong>Infinite Loop Conditions</strong>: Forgetting to increment indexing pointers or update loop state flags inside a while block.</li><li><strong>Off-By-One Indexing</strong>: Arrays are 0-indexed. Traversing up to index <code>arr.length</code> throws IndexOutOfBounds errors. Use <code>i < arr.length</code>.</li><li><strong>Ignoring Edge Cases</strong>: Forgetting to test empty arrays, zero inputs, duplicate elements, or singular linked list nodes. Always write validation guards!</li></ul>"
    },
    {
      "id": "dsa_pattern_1",
      "question": "1. Sliding Window Pattern",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">A highly optimized method for solving contiguous subarray or substring problems by maintaining a window that stretches or shrinks, reducing redundant $O(N^2)$ nested loops to a linear $O(N)$ traversal.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Linear structures (arrays, strings, or lists).</li><li>Asked to find contiguous subarrays, substrings, or sub-segments.</li><li>Problem seeks maximum, minimum, or unique characteristics within a segment size K or matching a target sum.</li></ul><strong>Code Implementation (Max Sum Subarray of Size K):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">sliding_window.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function maxSubarraySum(arr, k) {\n  let maxSum = 0, windowSum = 0;\n  // Initialize the first window of size k\n  for (let i = 0; i < k; i++) {\n    windowSum += arr[i];\n  }\n  maxSum = windowSum;\n  // Slide the window across the array\n  for (let i = k; i < arr.length; i++) {\n    windowSum += arr[i] - arr[i - k]; // Add next, remove previous\n    maxSum = Math.max(maxSum, windowSum);\n  }\n  return maxSum;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N)$ - Only a single linear pass is made across the array.</li><li><strong>Space Complexity:</strong> $O(1)$ - In-place calculation using constant extra space variables.</li></ul>"
    },
    {
      "id": "dsa_pattern_2",
      "question": "2. Two Pointers Pattern",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Uses two coordinate pointers traversing a collection at differing speeds or from opposing directions (e.g. left and right borders) to locate pairs or triplets satisfying specific mathematical limits.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Sorted linear collections (arrays, strings).</li><li>Searching for pairs, triplets, or unique indices summing to a target value.</li><li>Comparing elements from borders closing in toward the center.</li></ul><strong>Code Implementation (Two Sum Sorted Array):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">two_pointers.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function twoSumSorted(arr, target) {\n  let left = 0, right = arr.length - 1;\n  while (left < right) {\n    let sum = arr[left] + arr[right];\n    if (sum === target) return [left, right];\n    else if (sum < target) left++; // Need larger sum, shift right\n    else right--; // Need smaller sum, shift left\n  }\n  return [];\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N)$ - Traverses elements inward at most once.</li><li><strong>Space Complexity:</strong> $O(1)$ - Uses zero auxiliary space variables.</li></ul>"
    },
    {
      "id": "dsa_pattern_3",
      "question": "3. Fast & Slow Pointers (Hare & Tortoise)",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">A cyclic pointer algorithm that traverses a linear list at different intervals (slow pointer moves 1 step, fast pointer moves 2 steps) to detect loops or pinpoint mid-positions.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>LinkedList traversals or cyclic arrays.</li><li>Detecting loops, infinite state shifts, or cycle lengths.</li><li>Finding the exact middle node of a single LinkedList in a single pass.</li></ul><strong>Code Implementation (LinkedList Cycle Detection):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">fast_slow_pointer.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function hasCycle(head) {\n  let slow = head, fast = head;\n  while (fast !== null && fast.next !== null) {\n    slow = slow.next; // 1 step\n    fast = fast.next.next; // 2 steps\n    if (slow === fast) return true; // Loop cycle found!\n  }\n  return false;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N)$ - Linear search bounds prior to intersection.</li><li><strong>Space Complexity:</strong> $O(1)$ - Complete in-place loop cycle verification.</li></ul>"
    },
    {
      "id": "dsa_pattern_4",
      "question": "4. Merge Intervals Pattern",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Sorts dynamic coordinate intervals based on start boundaries to systematically unify overlapping ranges, resolving gaps and scheduling clashes.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Dynamic lists of intervals or numerical limits (e.g., schedules, coordinates).</li><li>Problems seeking to merge overlapping time slots or resolve meeting room clashes.</li><li>Checking if any schedules conflict or overlap.</li></ul><strong>Code Implementation (Merge Overlapping Intervals):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">merge_intervals.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function mergeIntervals(intervals) {\n  if (intervals.length <= 1) return intervals;\n  // Sort by starting boundaries\n  intervals.sort((a, b) => a[0] - b[0]);\n  let merged = [intervals[0]];\n  for (let i = 1; i < intervals.length; i++) {\n    let current = intervals[i];\n    let last = merged[merged.length - 1];\n    if (current[0] <= last[1]) {\n      last[1] = Math.max(last[1], current[1]); // Merge intervals\n    } else {\n      merged.push(current); // Disjoint interval\n    }\n  }\n  return merged;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N \\log N)$ - Sort takes $N \\log N$ while scanning takes linear time.</li><li><strong>Space Complexity:</strong> $O(N)$ - To store output list of merged intervals.</li></ul>"
    },
    {
      "id": "dsa_pattern_5",
      "question": "5. Tree Depth-First Search (DFS)",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Traverses deep branch structures recursively down to leaf boundaries first, using the runtime Call Stack to collect path metrics and heights.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Trees, hierarchical nodes, or graph traversals.</li><li>Problems asking for paths (e.g. node-to-leaf paths, path sum targets).</li><li>Checking depths, leaf nodes, or recursive ancestors.</li></ul><strong>Code Implementation (Path Sum Checking):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">tree_dfs.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function hasPathSum(root, sum) {\n  if (root === null) return false;\n  // Check if current node is a leaf node\n  if (root.left === null && root.right === null) {\n    return sum === root.val;\n  }\n  // Recursively search child paths\n  return hasPathSum(root.left, sum - root.val) || \n         hasPathSum(root.right, sum - root.val);\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N)$ - Visits each tree node exactly once in the worst case.</li><li><strong>Space Complexity:</strong> $O(H)$ - Recursion stack equal to tree height $H$.</li></ul>"
    },
    {
      "id": "dsa_pattern_6",
      "question": "6. Tree Breadth-First Search (BFS)",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Uses an explicit FIFO Queue to systematically visit and group tree structures layer-by-level (level-order traversal), locating closest node neighbors first.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Level-by-level outputs or grid layouts.</li><li>Shortest path or connection queries in unweighted structures.</li><li>Collecting child nodes at a specific depth layer.</li></ul><strong>Code Implementation (Level Order Traversal):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">tree_bfs.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function levelOrder(root) {\n  if (root === null) return [];\n  let result = [], queue = [root];\n  while (queue.length > 0) {\n    let levelSize = queue.length, currentLevel = [];\n    for (let i = 0; i < levelSize; i++) {\n      let node = queue.shift();\n      currentLevel.push(node.val);\n      if (node.left) queue.push(node.left);\n      if (node.right) queue.push(node.right);\n    }\n    result.push(currentLevel);\n  }\n  return result;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N)$ - Traverses each node exactly once.</li><li><strong>Space Complexity:</strong> $O(N)$ - FIFO Queue holds the largest layer depth layer width.</li></ul>"
    },
    {
      "id": "dsa_pattern_7",
      "question": "7. Subsets & Backtracking Pattern",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">A systematic DFS path generator that constructs possible sequences or combinations step-by-step, rolling back choices (backtracking) if a path hits dead ends.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Generating all permutations, combinations, or subsets.</li><li>Exhaustive path checking, grid solver mazes, or puzzle validators.</li><li>Problems seeking recursive combination trees.</li></ul><strong>Code Implementation (Generate Subsets):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">backtracking.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function generateSubsets(nums) {\n  let subsets = [];\n  function backtrack(start, current) {\n    subsets.push([...current]); // Record combination copy\n    for (let i = start; i < nums.length; i++) {\n      current.push(nums[i]);\n      backtrack(i + 1, current); // Recurse next elements\n      current.pop(); // Backtrack (remove element to try another)\n    }\n  }\n  backtrack(0, []);\n  return subsets;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(2^N)$ - Double choice paths at each element branch.</li><li><strong>Space Complexity:</strong> $O(N)$ - Max depth of the recursion tree.</li></ul>"
    },
    {
      "id": "dsa_pattern_8",
      "question": "8. Topological Sort (Kahn's Algorithm)",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Orders vertices of a Directed Acyclic Graph (DAG) linearly based on prereq dependencies by computing node in-degrees and using a queue to process resolved nodes.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Directed Acyclic Graphs (DAG).</li><li>Dependency scheduling problems (e.g. course planning, compiler build chains).</li><li>Determining cycles in directed graphs.</li></ul><strong>Code Implementation (Course Schedule Course Sorting):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">topological_sort.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function canFinishCourses(numCourses, prerequisites) {\n  let inDegree = Array(numCourses).fill(0);\n  let adjList = Array(numCourses).fill().map(() => []);\n  for (let [course, pre] of prerequisites) {\n    adjList[pre].push(course);\n    inDegree[course]++;\n  }\n  let queue = [];\n  for (let i = 0; i < numCourses; i++) {\n    if (inDegree[i] === 0) queue.push(i);\n  }\n  let count = 0;\n  while (queue.length > 0) {\n    let pre = queue.shift();\n    count++;\n    for (let course of adjList[pre]) {\n      inDegree[course]--;\n      if (inDegree[course] === 0) queue.push(course);\n    }\n  }\n  return count === numCourses;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(V + E)$ - Standard graph scanning over Vertices V and Edges E.</li><li><strong>Space Complexity:</strong> $O(V + E)$ - Storage for adjacency matrix representations.</li></ul>"
    },
    {
      "id": "dsa_pattern_9",
      "question": "9. Modified Binary Search",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">An efficient lookup pattern that halves the target search space at each iteration. It is widely used to solve search problems in sorted arrays, rotated arrays, or finding extreme value boundaries.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Sorted linear inputs (arrays or ranges).</li><li>Asked to achieve a logarithmic $O(\\log N)$ lookup performance.</li><li>Rotated sorted arrays or locating pivot boundaries.</li></ul><strong>Code Implementation (Search in Rotated Sorted Array):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">rotated_binary_search.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function searchRotated(arr, target) {\n  let left = 0, right = arr.length - 1;\n  while (left <= right) {\n    let mid = Math.floor((left + right) / 2);\n    if (arr[mid] === target) return mid;\n    \n    // Check if the left half is normally sorted\n    if (arr[left] <= arr[mid]) {\n      if (target >= arr[left] && target < arr[mid]) {\n        right = mid - 1; // Search left half\n      } else {\n        left = mid + 1;  // Search right half\n      }\n    } else { // Right half is normally sorted\n      if (target > arr[mid] && target <= arr[right]) {\n        left = mid + 1;  // Search right half\n      } else {\n        right = mid - 1; // Search left half\n      }\n    }\n  }\n  return -1;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(\\log N)$ - Array dimensions are split in half at each iteration step.</li><li><strong>Space Complexity:</strong> $O(1)$ - In-place iterations utilizing constant spaces.</li></ul>"
    },
    {
      "id": "dsa_pattern_10",
      "question": "10. Monotonic Stack Pattern",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Maintains a stack with strictly increasing or decreasing elements. As new indices arrive, it pops non-matching elements from the stack to locate boundaries in $O(N)$ linear time instead of $O(N^2)$ nested lookups.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Searching for the \"next greater element\", \"next smaller element\", or \"daily temperature wait times\".</li><li>Calculating largest rectangles in histogram charts.</li></ul><strong>Code Implementation (Next Greater Element):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">monotonic_stack.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function nextGreaterElement(arr) {\n  let result = Array(arr.length).fill(-1);\n  let stack = []; // Holds index references\n  for (let i = 0; i < arr.length; i++) {\n    // Pop indices from stack that have values smaller than current\n    while (stack.length > 0 && arr[stack[stack.length - 1]] < arr[i]) {\n      let idx = stack.pop();\n      result[idx] = arr[i]; // Found next greater element!\n    }\n    stack.push(i);\n  }\n  return result;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N)$ - Each element is pushed and popped from the stack at most once.</li><li><strong>Space Complexity:</strong> $O(N)$ - Auxiliary stack storage holds array references.</li></ul>"
    },
    {
      "id": "dsa_pattern_11",
      "question": "11. Trie (Prefix Tree) Pattern",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">An advanced search tree node structure representing character offsets, allowing rapid character lookups, string prefix matching, and spelling autocompletes in $O(L)$ time (where $L$ is word length).</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>String collections with massive dictionary overlaps.</li><li>Implementing search autocomplete engines.</li><li>Efficiently checking prefix validations (e.g. \"does word start with...\").</li></ul><strong>Code Implementation (Implement Trie Structure):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">trie_prefix_tree.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">class TrieNode {\n  constructor() {\n    this.children = {}; // Map: character -> TrieNode\n    this.isEndOfWord = false;\n  }\n}\n\nclass Trie {\n  constructor() {\n    this.root = new TrieNode();\n  }\n  // Inserts a word into the Trie\n  insert(word) {\n    let current = this.root;\n    for (let char of word) {\n      if (!current.children[char]) {\n        current.children[char] = new TrieNode();\n      }\n      current = current.children[char];\n    }\n    current.isEndOfWord = true;\n  }\n  // Returns true if the word is exactly in the Trie\n  search(word) {\n    let current = this.root;\n    for (let char of word) {\n      if (!current.children[char]) return false;\n      current = current.children[char];\n    }\n    return current.isEndOfWord;\n  }\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(L)$ - Lookups scale purely with word length $L$.</li><li><strong>Space Complexity:</strong> $O(AL)$ - Storage allocated for alphabetical connections.</li></ul>"
    },
    {
      "id": "dsa_pattern_12",
      "question": "12. Top 'K' Elements / Heap Pattern",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Utilizes a Min-Heap or Max-Heap structure to track the top largest or smallest $K$ elements from an unsorted collection or dynamic data stream in highly optimized $O(N \\log K)$ time.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Asked to find the \"Kth largest\" or \"Kth smallest\" elements.</li><li>Tracking top frequent elements or merging multiple pre-sorted arrays.</li></ul><strong>Code Implementation (Kth Largest Element):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">top_k_heap.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">// Simplified Min-Heap array-implementation for demonstration\nclass MinHeap {\n  constructor() { this.data = []; }\n  push(val) {\n    this.data.push(val);\n    this.data.sort((a,b) => a - b; // Keep sorted for min tracking\n  }\n  pop() { return this.data.shift(); }\n  peek() { return this.data[0]; }\n  size() { return this.data.length; }\n}\n\nfunction findKthLargest(nums, k) {\n  let minHeap = new MinHeap();\n  for (let num of nums) {\n    minHeap.push(num);\n    if (minHeap.size() > k) {\n      minHeap.pop(); // Evict smallest element, keeps K largest\n    }\n  }\n  return minHeap.peek();\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N \\log K)$ - Min heap is maintained at boundary height $K$ for all inputs.</li><li><strong>Space Complexity:</strong> $O(K)$ - Heap space restricted to maximum size $K$.</li></ul>"
    },
    {
      "id": "dsa_pattern_13",
      "question": "13. Dynamic Programming (0/1 Knapsack)",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">An optimization solver that breaks complex recursive paths into sub-problems, storing their outcomes (memoization/tabulation) to resolve overlapping calculation patterns in $O(N \\times C)$ boundaries.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Optimization challenges requesting maximum or minimum target results.</li><li>Decisions with overlapping choices (e.g. choosing whether to take or leave an index).</li><li>Dynamic subset sums or subset partitioned weights.</li></ul><strong>Code Implementation (Partition Equal Subset Sum):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">01_knapsack_dp.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function canPartition(nums) {\n  let totalSum = nums.reduce((a, b) => a + b, 0);\n  if (totalSum % 2 !== 0) return false; // Odd sum cannot partition\n  let target = totalSum / 2;\n  \n  let dp = Array(target + 1).fill(false);\n  dp[0] = true; // Base case: zero sum is always possible\n  \n  for (let num of nums) {\n    for (let j = target; j >= num; j--) {\n      dp[j] = dp[j] || dp[j - num]; // Take or Leave transition\n    }\n  }\n  return dp[target];\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(N \\times S)$ - Process array elements across total Target Sum $S$.</li><li><strong>Space Complexity:</strong> $O(S)$ - Tabular storage compressed to linear Sum boundaries.</li></ul>"
    },
    {
      "id": "dsa_pattern_14",
      "question": "14. Graph Depth-First Search (DFS)",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Explores fully along linear paths within graph nodes or grid coordinate matrices recursively, utilizing a visited set or boundary marking to identify closed connected segments.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Matrix representations representing maps, nodes, or pixel grids.</li><li>Asked to find \"number of connected islands\", closed clusters, or flood fills.</li></ul><strong>Code Implementation (Number of Islands Grid DFS):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">graph_dfs.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function numIslands(grid) {\n  if (!grid || grid.length === 0) return 0;\n  let count = 0;\n  \n  function dfs(r, c) {\n    // Boundary checks & check if cell is water/visited\n    if (r < 0 || c < 0 || r >= grid.length || c >= grid[0].length || grid[r][c] === '0') return;\n    grid[r][c] = '0'; // Mark as visited in-place!\n    dfs(r + 1, c); // Down\n    dfs(r - 1, c); // Up\n    dfs(r, c + 1); // Right\n    dfs(r, c - 1); // Left\n  }\n  \n  for (let r = 0; r < grid.length; r++) {\n    for (let c = 0; c < grid[0].length; c++) {\n      if (grid[r][c] === '1') {\n        count++; // Found a new island, explore it fully!\n        dfs(r, c);\n      }\n    }\n  }\n  return count;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(R \\times C)$ - Visits each coordinate row $R$ and column $C$ inside the matrix once.</li><li><strong>Space Complexity:</strong> $O(R \\times C)$ - Recursion stack in the worst case (all cells are land).</li></ul>"
    },
    {
      "id": "dsa_pattern_15",
      "question": "15. Shortest Path (Dijkstra's Algorithm)",
      "answer": "<strong>Concept Explanation:</strong><p style=\"margin: 0.5rem 0 1rem 0; color: var(--text-muted);\">Finds the minimum distance from a start vertex to all other vertices in a weighted graph by greedily expanding paths using a Priority Queue, checking connections incrementally.</p><strong>When to Recognize/Use:</strong><ul style=\"margin: 0.5rem 0 1rem 1.5rem; color: var(--text-muted); list-style-type: disc;\"><li>Weighted directed or undirected graphs.</li><li>Locating shortest path routes, minimum network delays, or cheapest flight path schedules.</li></ul><strong>Code Implementation (Dijkstra's Shortest Path):</strong><div class=\"console-box\" style=\"margin: 0.5rem 0 1rem 0;\"><div class=\"console-header\" style=\"padding: 0.4rem 1rem;\"><div style=\"font-size: 0.75rem; font-family: var(--font-mono); color: var(--accent-cyan);\">dijkstras_shortest_path.js</div></div><pre style=\"margin: 0; padding: 1rem; overflow-x: auto; font-family: var(--font-mono); font-size: 0.85rem; color: #adbac7; line-height: 1.5; background: #06050b;\">function dijkstra(graph, startNode, numNodes) {\n  let distances = Array(numNodes).fill(Infinity);\n  distances[startNode] = 0;\n  let visited = Array(numNodes).fill(false);\n  \n  for (let i = 0; i < numNodes - 1; i++) {\n    // Find the minimum distance node not yet visited\n    let minDistance = Infinity, u = -1;\n    for (let v = 0; v < numNodes; v++) {\n      if (!visited[v] && distances[v] <= minDistance) {\n        minDistance = distances[v];\n        u = v;\n      }\n    }\n    if (u === -1) break;\n    visited[u] = true;\n    \n    // Update neighbors distances\n    for (let neighbor = 0; neighbor < numNodes; neighbor++) {\n      let weight = graph[u][neighbor];\n      if (weight > 0 && !visited[neighbor]) {\n        let newDistance = distances[u] + weight;\n        if (newDistance < distances[neighbor]) {\n          distances[neighbor] = newDistance;\n        }\n      }\n    }\n  }\n  return distances;\n}</pre></div><strong>Complexity Analysis:</strong><ul style=\"margin: 0.5rem 0 0 1.5rem; color: var(--text-muted); list-style-type: square;\"><li><strong>Time Complexity:</strong> $O(V^2)$ - Standard Dijkstra using array minimum lookups (optimizable to $O(E \\log V)$ using Min-Priority Queues).</li><li><strong>Space Complexity:</strong> $O(V)$ - Storage array variables tracking visited flags and distance indices.</li></ul>"
    }
  ],
  "System Design": [
    {
      "id": "sys_1",
      "question": "What is the CAP Theorem?",
      "answer": "The CAP Theorem states that a distributed data store can simultaneously provide at most two of the following three guarantees:\n1. **Consistency (C)**: Every read receives the most recent write or an error.\n2. **Availability (A)**: Every non-failing node returns a non-error response (without guarantee that it contains the most recent write).\n3. **Partition Tolerance (P)**: The system continues to operate despite arbitrary message loss or network partition splits. Under a partition, you must choose between Consistency (CP) or Availability (AP)."
    },
    {
      "id": "sys_2",
      "question": "What are the common caching patterns (Cache-Aside, Write-Through, Write-Behind)?",
      "answer": "- **Cache-Aside (Lazy Loading)**: The application checks the cache first. If a hit occurs, it returns data. If a miss occurs, it queries the database, returns data, and populates the cache for future queries.\n- **Write-Through**: The application writes directly to the cache, and the cache synchronously writes to the database. Ensures data is always fresh, but incurs write latency.\n- **Write-Behind (Write-Back)**: The application writes to the cache, which acknowledges immediately. The cache then asynchronously writes back to the database in batches, maximizing write performance but risking data loss in power failure."
    }
  ],
  "RAG, MCP, Vector Database, LLM, Tokens": [
    {
      "id": "rag_1",
      "question": "What is Retrieval-Augmented Generation (RAG)?",
      "answer": "Retrieval-Augmented Generation (RAG) is an architectural pattern that enhances LLM generation by retrieving relevant information from external knowledge bases (e.g., vectorized documents) and appending it to the LLM's prompt. This allows the model to reference up-to-date, proprietary data and cite specific sources while reducing factual inaccuracies (hallucinations)."
    },
    {
      "id": "rag_2",
      "question": "Explain Vector Databases and how similarity search works.",
      "answer": "Vector Databases (like Pinecone, Milvus, Chroma) are built to store and index high-dimensional vector embeddings generated by machine learning models. Unlike traditional relational databases, they perform highly optimized mathematical similarity searches (e.g., Cosine Similarity, Dot Product, or Euclidean Distance) using algorithms like Hierarchical Navigable Small World (HNSW) to locate the closest conceptual vectors in milliseconds."
    },
    {
      "id": "rag_3",
      "question": "What is Model Context Protocol (MCP)?",
      "answer": "The Model Context Protocol (MCP) is an open standard that allows developers to build secure, standardized connections between AI models and their data sources or tools. Rather than writing custom API wrappers for every integration, MCP defines a unified client-server architecture where models can discover, inspect, and interact with files, databases, and APIs in a sandboxed, uniform way."
    }
  ],
  "Langraph, MultiAgent, Crew AI": [
    {
      "id": "agent_1",
      "question": "What is a Multi-Agent System and how does it differ from single LLM execution?",
      "answer": "A Multi-Agent System breaks down a complex objective into modular tasks managed by specialized AI persona 'agents' (e.g., Researcher, Writer, Code Reviewer). While a single LLM tries to solve everything in one shot, multi-agent networks use orchestration frameworks to share context, loop on feedback, and use external tools in a cyclic, cooperative fashion, vastly improving accuracy and handling complex multi-stage workflows."
    },
    {
      "id": "agent_2",
      "question": "Explain State Management in LangGraph.",
      "answer": "LangGraph is designed for building stateful, multi-actor applications with LLMs by modeling workflows as graphs (nodes as agents/tools, and edges as transition logic). It features robust state management where a global state object is passed from node to node, allowing cyclic loops, human-in-the-loop approvals, and full thread persistence to pause, resume, or replay agent steps safely."
    }
  ]
};

// Fallback high-quality practice challenges
const FALLBACK_PRACTICE_DATA = {
  complexity: [
    {
      topic: "DSA: Array Search",
      code: `function findElement(arr, target) {\n  for (let i = 0; i < arr.length; i++) {\n    if (arr[i] === target) return i;\n  }\n  return -1;\n}`,
      options: ["O(1)", "O(log n)", "O(n)", "O(n^2)"],
      answer: 2,
      explanation: "The algorithm traverses the array sequentially. In the worst-case scenario (target is at the very end or not present), it must check all 'n' elements, yielding a linear time complexity of O(n)."
    },
    {
      topic: "DSA: Binary Search",
      code: `function binarySearch(arr, target) {\n  let left = 0;\n  let right = arr.length - 1;\n  while (left <= right) {\n    let mid = Math.floor((left + right) / 2);\n    if (arr[mid] === target) return mid;\n    else if (arr[mid] < target) left = mid + 1;\n    else right = mid - 1;\n  }\n  return -1;\n}`,
      options: ["O(1)", "O(log n)", "O(n)", "O(n log n)"],
      answer: 1,
      explanation: "In each step, the algorithm divides the search space in half. The number of iterations grows logarithmically as the array size increases, resulting in an O(log n) time complexity."
    },
    {
      topic: "DSA: Matrix Multiplication",
      code: `function multiplyMatrices(A, B) {\n  let n = A.length;\n  let C = Array(n).fill().map(() => Array(n).fill(0));\n  for (let i = 0; i < n; i++) {\n    for (let j = 0; j < n; j++) {\n      for (let k = 0; k < n; k++) {\n        C[i][j] += A[i][k] * B[k][j];\n      }\n    }\n  }\n  return C;\n}`,
      options: ["O(n)", "O(n log n)", "O(n^2)", "O(n^3)"],
      answer: 3,
      explanation: "There are three nested loops, each iterating up to 'n' (the dimension of the matrices). The total operations executed scale cubically, making it O(n^3)."
    },
    {
      topic: "DSA: Fibonacci Recursion",
      code: `function fib(n) {\n  if (n <= 1) return n;\n  return fib(n - 1) + fib(n - 2);\n}`,
      options: ["O(log n)", "O(n)", "O(n^2)", "O(2^n)"],
      answer: 3,
      explanation: "Each call to fib(n) spawns two additional recursive calls (fib(n-1) and fib(n-2)). The recursion tree doubles in size at each level of depth, resulting in an exponential time complexity of O(2^n)."
    }
  ],
  scenarios: {
    "System Design": [
      {
        subtype: "happy_path",
        title: "Design a High-Traffic Notification System",
        description: "Your system must process and dispatch up to 100,000 notifications per second across multiple channels (Email, SMS, Push alerts) with minimal latency.",
        prompt: "Design the **happy path** for this notification service. Explain what components are involved from the initial client request to successful delivery across the network, and how you ensure high throughput."
      },
      {
        subtype: "failure_path",
        title: "Mitigate Notification Queue Bottlenecks",
        description: "During a major marketing campaign, the SMS gateway provider suddenly starts rate-limiting your outgoing requests, causing a massive backlog that delays critical system emails.",
        prompt: "Explain how you will **handle and fix this failure path**. What caching, queuing, or circuit breaker mechanisms would you implement to isolate SMS delivery issues and ensure emails still go out instantly?"
      }
    ],
    "RAG, MCP, Vector Database, LLM, Tokens": [
      {
        subtype: "happy_path",
        title: "Design a Real-time Multi-Source RAG Sync",
        description: "Design an enterprise internal search system that ingests updates from Slack, Google Drive, and Confluence, generates vector embeddings, and updates the knowledge bank in near real-time.",
        prompt: "Detail your **happy path** pipeline. Explain how you partition the data, generate chunk embeddings, and index them in a vector database for semantic retrieval during user queries."
      },
      {
        subtype: "failure_path",
        title: "Fix Token Cost Spikes and Context Bloat",
        description: "A sudden influx of long documents causes your LLM token usage to skyrocket, leading to high bills and context window overflows that make the search results highly inaccurate.",
        prompt: "Detail your **failure path recovery and optimization**. How would you handle and fix this issue? Consider chunking strategies, reranking algorithms, local caching, or token limits."
      }
    ],
    "Langraph, MultiAgent, Crew AI": [
      {
        subtype: "happy_path",
        title: "Design a Cooperative Agent Content Ingestion Pipeline",
        description: "Design a system where three agents cooperate: Agent A parses raw web documents, Agent B extracts structured business facts, and Agent C writes an executive summary.",
        prompt: "Design the **happy path** graph flow. How do the agents hand off state, share a mutual database, and validate that summaries are factually accurate without human intervention?"
      },
      {
        subtype: "failure_path",
        title: "Resolve Infinite Tool-Calling Loops",
        description: "Agent B encounters a document with confusing formatting, leading it to call a lookup tool repeatedly in an infinite cyclic loop, consuming massive amounts of API usage without ever completing the task.",
        prompt: "How do you **handle and fix this failure path**? What state limits, token counters, validation guards, or human-in-the-loop check stops would you introduce in your LangGraph workflow?"
      }
    ]
  }
};

// ----------------------------------------------------
// UTILITY FUNCTIONS
// ----------------------------------------------------

// Parses the Topics.xlsx file to inspect structure
function getExcelTopics() {
  if (!fs.existsSync(EXCEL_PATH)) {
    return { error: 'Topics.xlsx not found in workspace' };
  }
  try {
    const workbook = xlsx.readFile(EXCEL_PATH);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    
    // Format rows nicely to send to UI
    const rows = [];
    data.forEach((row, i) => {
      if (row.length > 0) {
        rows.push({
          rowNum: i + 1,
          colB: row[1] || '',
          colC: row[2] || ''
        });
      }
    });
    return { success: true, sheetName, totalRows: data.length, rows };
  } catch (err) {
    return { error: 'Failed parsing excel: ' + err.message };
  }
}

// Read local DB
function readDb() {
  try {
    if (!fs.existsSync(DB_PATH)) {
      // Return default empty state
      return {
        progress: {
          "AI": { completedCards: [], totalQuestions: 2 },
          "ML": { completedCards: [], totalQuestions: 2 },
          "NLP": { completedCards: [], totalQuestions: 2 },
          "DSA": { completedCards: [], complexityCorrect: 0, complexityTotal: 0, totalQuestions: 3 },
          "System Design": { completedCards: [], scenariosAttempted: 0, totalQuestions: 2 },
          "RAG, MCP, Vector Database, LLM, Tokens": { completedCards: [], scenariosAttempted: 0, totalQuestions: 3 },
          "Langraph, MultiAgent, Crew AI": { completedCards: [], scenariosAttempted: 0, totalQuestions: 2 }
        },
        scenariosHistory: [],
        complexityHistory: []
      };
    }
    const raw = fs.readFileSync(DB_PATH, 'utf8');
    return JSON.parse(raw);
  } catch (err) {
    console.error('Error reading journey db:', err);
    return {};
  }
}

// Write to local DB
function writeDb(data) {
  try {
    fs.writeFileSync(DB_PATH, JSON.stringify(data, null, 2), 'utf8');
    return true;
  } catch (err) {
    console.error('Error writing journey db:', err);
    return false;
  }
}

// Initialize Gemini Client
function getGeminiClient(apiKey) {
  const key = apiKey || process.env.GEMINI_API_KEY;
  if (!key) return null;
  return new GoogleGenAI({ apiKey: key });
}

// ----------------------------------------------------
// EXPRESS API ENDPOINTS
// ----------------------------------------------------

// Get Excel structures & curated static topics
app.get('/api/topics', (req, res) => {
  const excelInfo = getExcelTopics();
  res.json({
    learningBank: DEFAULT_LEARNING_BANK,
    excel: excelInfo
  });
});

// Get user progress
app.get('/api/journey', (req, res) => {
  res.json(readDb());
});

// Update card completion
app.post('/api/journey/complete-card', (req, res) => {
  const { topic, cardId } = req.body;
  if (!topic || !cardId) {
    return res.status(400).json({ error: 'Missing topic or cardId' });
  }

  const db = readDb();
  if (!db.progress[topic]) {
    db.progress[topic] = { completedCards: [] };
  }

  if (!db.progress[topic].completedCards.includes(cardId)) {
    db.progress[topic].completedCards.push(cardId);
  }

  writeDb(db);
  res.json({ success: true, db });
});

// Generate dynamic question using Gemini (or fallback demo mode)
app.post('/api/generate-challenge', async (req, res) => {
  const { topic, type, apiKey } = req.body;
  if (!topic || !type) {
    return res.status(400).json({ error: 'Missing topic or type' });
  }

  const gemini = getGeminiClient(apiKey);

  if (!gemini) {
    // RUN IN DEMO / FALLBACK MODE
    console.log(`[API] Gemini API key not found. Using local mock generator for topic: ${topic}`);
    if (type === 'complexity') {
      const idx = Math.floor(Math.random() * FALLBACK_PRACTICE_DATA.complexity.length);
      const question = FALLBACK_PRACTICE_DATA.complexity[idx];
      return res.json({ mode: 'mock', data: question });
    } else {
      const items = FALLBACK_PRACTICE_DATA.scenarios[topic] || FALLBACK_PRACTICE_DATA.scenarios["System Design"];
      const idx = Math.floor(Math.random() * items.length);
      const question = items[idx];
      return res.json({ mode: 'mock', data: question });
    }
  }

  try {
    if (type === 'complexity') {
      // Ask Gemini to generate a code snippet and multiple choice complexity question
      const prompt = `Generate a programming code snippet in JavaScript, Python, or Go related to the topic of '${topic}'. 
Provide 4 multiple choice options for its worst-case Time Complexity (e.g. O(1), O(log n), O(n), O(n log n), O(n^2), O(2^n)).
Return the result strictly as a valid JSON object with the following keys and no extra formatting or markdown blocks:
{
  "topic": "String representing specific DSA subtopic",
  "code": "String containing the code snippet with proper formatting and newlines",
  "options": ["Option 1", "Option 2", "Option 3", "Option 4"],
  "answer": Integer (0 to 3 index of the correct option in the options array),
  "explanation": "Detailed explanation explaining why it has that complexity"
}`;

      const ai = gemini.models.get({ model: 'gemini-1.5-flash' });
      const response = await ai.generateContent({
        contents: prompt,
        generationConfig: { responseMimeType: "application/json" }
      });
      const parsed = JSON.parse(response.text.trim());
      return res.json({ mode: 'live', data: parsed });
    } else {
      // System Design scenario: Randomly decide happy or failure path
      const subType = Math.random() > 0.5 ? 'happy_path' : 'failure_path';
      let prompt = '';
      if (subType === 'happy_path') {
        prompt = `Generate a modern system design scenario about '${topic}'. It should focus on a **Happy Path** setup where a specific high-performance or complex workflow needs to be designed from scratch.
Return the result strictly as a valid JSON object with the following keys and no extra formatting:
{
  "subtype": "happy_path",
  "title": "A short descriptive title for the challenge",
  "description": "A comprehensive description of the happy path scenario including metrics, goals, and core system parameters.",
  "prompt": "Ask the user to explain how they would design this architecture from scratch, step-by-step, including what components they would use."
}`;
      } else {
        prompt = `Generate a modern system design scenario about '${topic}'. It should focus on a **Failure Path** setup where a concrete, specific error, bottleneck, outage, rate limit, or sync issue has occurred in an existing setup.
Return the result strictly as a valid JSON object with the following keys and no extra formatting:
{
  "subtype": "failure_path",
  "title": "A short descriptive title for the failure incident",
  "description": "Describe the outage, bottleneck, or failure scenario in details (e.g. what service is failing, what errors are returned, what rate limits are hit).",
  "prompt": "Ask the user to explain how they would handle, mitigate, and fix this specific failure path to ensure system resiliency."
}`;
      }

      const ai = gemini.models.get({ model: 'gemini-1.5-flash' });
      const response = await ai.generateContent({
        contents: prompt,
        generationConfig: { responseMimeType: "application/json" }
      });
      const parsed = JSON.parse(response.text.trim());
      return res.json({ mode: 'live', data: parsed });
    }
  } catch (err) {
    console.error('Gemini call failed, falling back to mock data:', err);
    // Graceful fallback
    if (type === 'complexity') {
      const idx = Math.floor(Math.random() * FALLBACK_PRACTICE_DATA.complexity.length);
      return res.json({ mode: 'fallback-mock', data: FALLBACK_PRACTICE_DATA.complexity[idx], error: err.message });
    } else {
      const items = FALLBACK_PRACTICE_DATA.scenarios[topic] || FALLBACK_PRACTICE_DATA.scenarios["System Design"];
      const idx = Math.floor(Math.random() * items.length);
      return res.json({ mode: 'fallback-mock', data: items[idx], error: err.message });
    }
  }
});

// Evaluate Scenario Submission
app.post('/api/evaluate-scenario', async (req, res) => {
  const { challenge, answer, apiKey } = req.body;
  if (!challenge || !answer) {
    return res.status(400).json({ error: 'Missing challenge or answer' });
  }

  const gemini = getGeminiClient(apiKey);
  
  // Local Database logging setup
  const db = readDb();
  const historyItem = {
    timestamp: new Date().toISOString(),
    title: challenge.title,
    subtype: challenge.subtype,
    description: challenge.description,
    userAnswer: answer,
    feedback: ""
  };

  if (!gemini) {
    // Local smart heuristic analyzer
    console.log('[API] Evaluating using local heuristic keywords analyzer.');
    let grade = 'B';
    let summary = 'A very thoughtful response!';
    let criticalReview = '';
    const answerLower = answer.toLowerCase();

    // Look for architectural terms
    const keywords = ['fallback', 'cache', 'caching', 'retry', 'circuit breaker', 'queue', 'redis', 'kafka', 'rate limit', 'load balancer', 'redundancy', 'monitoring'];
    const matched = keywords.filter(w => answerLower.includes(w));
    
    if (challenge.subtype === 'happy_path') {
      if (answer.length < 80) {
        grade = 'C';
        summary = 'Your design is a bit too brief. In real system design interviews, you need to detail component interfaces and data flow.';
        criticalReview = 'Add description of load balancers, caching layers, and database choices to make this robust.';
      } else if (matched.length >= 3) {
        grade = 'A';
        summary = 'Excellent design! You successfully detailed a structured data path and incorporated key scaling components.';
        criticalReview = `Great job including key concepts: ${matched.join(', ')}. Keep focusing on data model layout.`;
      } else {
        grade = 'B';
        summary = 'A solid foundational setup. To elevate it, focus more on data ingestion pathways.';
        criticalReview = 'Consider explaining how scale affects your database writes and whether messaging queues would help.';
      }
    } else {
      // Failure path analysis
      if (answer.length < 80) {
        grade = 'C';
        summary = 'The explanation is too short. Resilient architectures require solid redundancy layouts.';
        criticalReview = 'Mention specific patterns like Exponential Backoff, Dead Letter Queues, or Circuit Breakers to handle this.';
      } else if (matched.includes('circuit breaker') || matched.includes('fallback') || matched.includes('queue') || matched.includes('retry')) {
        grade = 'A';
        summary = 'Extremely resilient fix! You utilized proper failover patterns to isolate component degradation.';
        criticalReview = `Excellent use of resilient patterns: ${matched.join(', ')}. This fully secures the system from cascading failure.`;
      } else {
        grade = 'B';
        summary = 'Good general troubleshooting steps. To make it enterprise-grade, build automated self-healing mechanisms.';
        criticalReview = 'Introduce structured queues (like RabbitMQ/SQS) and circuit breakers to automatically throttle connections.';
      }
    }

    const localFeedback = { grade, summary, criticalReview };
    historyItem.feedback = localFeedback;
    db.scenariosHistory.push(historyItem);

    // Update section metrics
    const sectionName = challenge.topic || "System Design";
    if (db.progress[sectionName]) {
      db.progress[sectionName].scenariosAttempted = (db.progress[sectionName].scenariosAttempted || 0) + 1;
    }
    writeDb(db);

    return res.json({ mode: 'mock', evaluation: localFeedback, db });
  }

  try {
    const prompt = `You are an expert Principal Systems Architect grading a system design interview candidate.
Challenge Title: "${challenge.title}"
Challenge Description: "${challenge.description}"
Challenge Subtype: "${challenge.subtype}"
Question Prompt: "${challenge.prompt}"

Candidate's Answer:
"${answer}"

Evaluate the answer. Focus heavily on:
1. Did they correctly address the specific path? (Happy Path: correct component structure, scalability. Failure Path: proper recovery, isolation, self-healing).
2. Grade them with a single character letter grade ('A', 'B', 'C', or 'F').
3. Provide a constructive summary and a critical review of what is missing or how to improve.

Return the evaluation strictly as a valid JSON object matching this schema with no extra text or markdown formatting:
{
  "grade": "A/B/C/F",
  "summary": "High-level summary of their design pros",
  "criticalReview": "Detailed analysis of gaps, failure vulnerabilities, and architectural recommendations"
}`;

    const ai = gemini.models.get({ model: 'gemini-1.5-flash' });
    const response = await ai.generateContent({
      contents: prompt,
      generationConfig: { responseMimeType: "application/json" }
    });
    const parsed = JSON.parse(response.text.trim());
    
    historyItem.feedback = parsed;
    db.scenariosHistory.push(historyItem);

    // Update metrics
    const sectionName = challenge.topic || "System Design";
    if (db.progress[sectionName]) {
      db.progress[sectionName].scenariosAttempted = (db.progress[sectionName].scenariosAttempted || 0) + 1;
    }
    writeDb(db);

    return res.json({ mode: 'live', evaluation: parsed, db });
  } catch (err) {
    console.error('Gemini evaluation failed, returning smart heuristic analysis:', err);
    // Simple fallback logic
    const localFeedback = {
      grade: 'B',
      summary: 'Heuristic Review: Good initial conceptual design layout.',
      criticalReview: `Architectural analysis tool: Your response contains ${answer.split(' ').length} words. To improve scalability, make sure you explicitly mention asynchronous messaging, decoupling layers, and backoff retries.`
    };
    historyItem.feedback = localFeedback;
    db.scenariosHistory.push(historyItem);
    writeDb(db);
    return res.json({ mode: 'fallback-mock', evaluation: localFeedback, db, error: err.message });
  }
});

// Update DSA complexity scorecard
app.post('/api/journey/update-complexity', (req, res) => {
  const { isCorrect, selectedOption, correctOption, questionTopic } = req.body;
  
  const db = readDb();
  if (!db.progress["DSA"]) {
    db.progress["DSA"] = { completedCards: [], complexityCorrect: 0, complexityTotal: 0 };
  }
  
  db.progress["DSA"].complexityTotal = (db.progress["DSA"].complexityTotal || 0) + 1;
  if (isCorrect) {
    db.progress["DSA"].complexityCorrect = (db.progress["DSA"].complexityCorrect || 0) + 1;
  }

  db.complexityHistory.push({
    timestamp: new Date().toISOString(),
    topic: questionTopic,
    isCorrect,
    selectedOption,
    correctOption
  });

  writeDb(db);
  res.json({ success: true, db });
});

// ----------------------------------------------------
// RUN SERVER
// ----------------------------------------------------
app.listen(PORT, () => {
  console.log(`====================================================`);
  console.log(`🚀 Premium Adaptive Learning Server Running!`);
  console.log(`🔗 Interface: http://localhost:${PORT}`);
  console.log(`📅 Local time: ${new Date().toLocaleString()}`);
  console.log(`====================================================`);
  
  // Test excel loading on startup
  const xlTest = getExcelTopics();
  if (xlTest.error) {
    console.log(`⚠️ Excel Warning: ${xlTest.error}`);
  } else {
    console.log(`📊 Successfully parsed Topics.xlsx from workspace! Sheet: ${xlTest.sheetName}, Rows: ${xlTest.totalRows}`);
  }
});
