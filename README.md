# detailler_pgn
Tool to convert san moves to enhanced long algebraic moves<p>

Prerequisites :<br>
copy [nettoyer_epd.exe](https://github.com/chris13300/detailler_pgn/blob/main/detailler_pgn/bin/Debug/nettoyer_epd.exe) to your DETAILLER_PGN folder<br>
copy [pgn-extract.exe](https://github.com/chris13300/detailler_pgn/blob/main/detailler_pgn/bin/Debug/pgn-extract.exe) to your DETAILLER_PGN folder<p>
command : detailler_pgn.exe path_to_your_pgn_file.pgn

# How it works ?
- From a PGN file with commented/annotated moves :<br>
![commented_pgn](https://github.com/chris13300/detailler_pgn/blob/main/detailler_pgn/bin/Debug/commented_pgn.jpg)<p>

- we use SCID to clean most comments and annotations :<br>
![scid](https://github.com/chris13300/detailler_pgn/blob/main/detailler_pgn/bin/Debug/scid.jpg)<br>
![1_uncommented_pgn](https://github.com/chris13300/detailler_pgn/blob/main/detailler_pgn/bin/Debug/1_uncommented_pgn.jpg)<p>

- detailler_pgn cleans most optional tags :<br>
![2_headers_cleaned](https://github.com/chris13300/detailler_pgn/blob/main/detailler_pgn/bin/Debug/2_headers_cleaned.jpg)<p>

- detailler_pgn use pgn-extract to get the long algebraic moves (ex : g2g4) :<br>
![3_pgn-extract](https://github.com/chris13300/detailler_pgn/blob/main/detailler_pgn/bin/Debug/3_pgn-extract.jpg)<p>

- finally we get something like this :<br>
![4_detailled_moves](https://github.com/chris13300/detailler_pgn/blob/main/detailler_pgn/bin/Debug/4_detailled_moves.jpg)<p>
