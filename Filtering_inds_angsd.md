# To do some basic filtering of minimum individuals for beagle files
```
# For minimum individuals of 180:
zcat your_mafs_file.gz | awk '$8 >= 180 {print $0}' > minInd180.txt

cut -f1,2 minInd180.txt | awk '{print $1"_"$2}' > markers.txt
{ zcat your_zipped_beagle.gz | head -n 1 && zgrep -Ff <(tail -n +2 markers.txt) your_zipped_beagle.gz; } > new_filtered_minInd180.beagle # Note that you will have to gzip this back up or pipe it to gzip
```
