chmod +x lawsuit_generator.py
./lawsuit_generator.py \
    --firm_name "Your Law Firm" \
    --case "Case Title Here" \
    --output "lawsuit.pdf" \
    --file "lawsuit_body.txt" \
    --exhibits "caption1.txt" "image1.png" "caption2.txt" "image2.jpg" \
    --index "index.pdf" \
    --pickle "lawsuit.pickle"