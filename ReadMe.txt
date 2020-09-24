This is an object-oriented Sudoku Solver combining the "cross hatching" strategy with a backtrack algorithm. It solves easy puzzles in about 300 µsecs and really hard ones usually in under 150 msecs (when compiled) on my old AMD Athlon XP1800+. The most evil one I could find is in the screen shot ~~ Download is 23.6 kB. 

Verion 1
Update #1: Added Timeout and a rudimentary Contradiction Check. 

Update #2: Now chooses direction of backtrack (down or up) and better Navigation. 

Update #3: Added Animation mode. 

Update #4: Added File handling and a better Contradiction Detection.

Update #5: Added Vox and fixed some minor quirks. 
           Improved solving the puzzle grid as far as it goes before entering backtrack.
           Added commandline params.         

Update #6: Killed a strange bug with terminating while solving and added printing.

Version 2
Update #1: Changed recursion to prefer cells with the least number of possible values (tnx to 
           Derio from whose submission I got this idea). This has improved speed tremenduously.

Update #2: Added a proper About Box
           Added killer heuristic (which got the timing down by up to 70%)
           Some other speed improvements.
           Deleted Monopolize - it's no longer necessary.
           Added Animate Speed Selection.


Note I: 
SKP (Sudoku Puzzle) files are 81 characters long and contain one puzzle; cells containing a number are indicated by the appropiate figure and free cells are denoted by any non-numeric character - the first character being for the top left cell (A:1) and the last character for the bottom right cell (I:9). They may be edited with the standard Windows editor.

Note II:
If you want to use this program to assist you in solving puzzles manually, click Hide before clicking Solve. When the solution is ready you can either click on a solution square to reveal it's content or click Hide again to reveal them all. 

Note III:
If you assign .SKP files to Sudoku.EXE (right click on a .SKP file, Open With, Always) then you can also doubleclick on .SKP files to run Sudoku.EXE.

PS
Load NotSoEval.SKP and try to figure out which number belongs into the first cell and type it in. You will see that the puzzle is solvable just by (simple) logic and without any backtracking. 