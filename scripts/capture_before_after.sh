#!/bin/bash

# Capture before/after screenshots of font comparison presentations

BEFORE_ID="1otnD0T4WPmX4wRoCKlpNjRkFJ99CIAfvDz9eguJG5RI"
AFTER_ID="17llqEKwQJdKa26v82dC0jcwMMheuF8IHJp1iEa6TF2Y"

BEFORE_URL="https://docs.google.com/presentation/d/${BEFORE_ID}/pub"
AFTER_URL="https://docs.google.com/presentation/d/${AFTER_ID}/pub"

echo "🎯 CAPTURING BEFORE/AFTER FONT COMPARISON SCREENSHOTS"
echo "===================================================="
echo ""

# Create screenshots directory
mkdir -p screenshots

echo "📸 STEP 1: Capturing BEFORE presentation (default fonts)"
echo "🔗 Opening: $BEFORE_URL"
open "$BEFORE_URL"

echo "⏳ Waiting 8 seconds for presentation to load..."
sleep 8

echo "📸 Taking BEFORE screenshot..."
screencapture -x "screenshots/BEFORE_default_fonts.png"
echo "✅ BEFORE screenshot saved: screenshots/BEFORE_default_fonts.png"

echo ""
echo "📸 STEP 2: Capturing AFTER presentation (Merriweather/Inter fonts)"
echo "🔗 Opening: $AFTER_URL"
open "$AFTER_URL"

echo "⏳ Waiting 8 seconds for presentation to load..."
sleep 8

echo "📸 Taking AFTER screenshot..."
screencapture -x "screenshots/AFTER_merriweather_inter_fonts.png"
echo "✅ AFTER screenshot saved: screenshots/AFTER_merriweather_inter_fonts.png"

echo ""
echo "🎯 BEFORE/AFTER COMPARISON COMPLETE!"
echo "===================================="
echo "📁 Screenshots saved in screenshots/ directory:"
echo "  📄 BEFORE: screenshots/BEFORE_default_fonts.png"
echo "  📄 AFTER: screenshots/AFTER_merriweather_inter_fonts.png"
echo ""
echo "🔍 COMPARISON DETAILS:"
echo "📋 BEFORE Presentation (Default Fonts):"
echo "  🆔 ID: $BEFORE_ID"
echo "  📝 URL: $BEFORE_URL"
echo ""
echo "📋 AFTER Presentation (Merriweather/Inter):"
echo "  🆔 ID: $AFTER_ID" 
echo "  📝 URL: $AFTER_URL"
echo ""
echo "✅ Font pair change demonstration complete!"