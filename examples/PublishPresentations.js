/**
 * PUBLISH PRESENTATIONS FOR SCREENSHOT ACCESS
 * Make acid test presentations publicly viewable
 */

function publishAcidTestPresentations() {
  console.log('📤 Publishing acid test presentations for screenshot access...');
  
  // Presentation IDs from acid test
  const presentationIds = [
    '1sfTWcaHlFBfCbejitqI3z_7XufA1Lkqc50B5VlR7Nqo', // Combination 1
    '1o5N0RyzGr6CkERXXpWK7w9n18XhyHxd-64BQTuMHOZM', // Combination 2
    '1JgGQ4F3-npYjuGb0Kn6nkloPHZmfbMdCVhI1E3VHOXs', // Combination 3
    '1CF-GDQtkyrmCzj8zSteAPr37mzs_Wp-C2-z2wZrGpNQ', // Combination 4
    '1Xw4cCWIzEaD02sx67Lhe5fiuCEODF1YXZCST6fta3f4', // Combination 5
    '1_0UCt5x75hxV5EbERMDsi_ubfe71DOn7BphRVabp1J8', // Combination 6
    '1yaDurSgpOz2yQ66xXaxSuiAJPFKw7wSzVtsQpTNU4nM', // Combination 7
    '1yI27WeRtdWvJCqsx-Fp6EXgPba6fh1ikubYwaz1oyiw'  // Combination 8
  ];
  
  const results = [];
  
  try {
    for (let i = 0; i < presentationIds.length; i++) {
      const presentationId = presentationIds[i];
      console.log(`\n📋 Publishing presentation ${i + 1}: ${presentationId}`);
      
      try {
        // Method 1: Make file publicly viewable via Drive API
        const driveUrl = `https://www.googleapis.com/drive/v3/files/${presentationId}/permissions`;
        
        const permissionPayload = {
          role: 'reader',
          type: 'anyone'
        };
        
        const response = UrlFetchApp.fetch(driveUrl, {
          method: 'POST',
          headers: {
            'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify(permissionPayload)
        });
        
        if (response.getResponseCode() === 200) {
          console.log(`  ✅ Made publicly viewable via Drive API`);
        } else {
          console.log(`  ⚠️ Drive API response: ${response.getResponseCode()}`);
        }
        
        // Method 2: Try to publish via Slides API
        try {
          const slidesUrl = `https://slides.googleapis.com/v1/presentations/${presentationId}:batchUpdate`;
          
          const batchUpdatePayload = {
            requests: [{
              updatePresentationProperties: {
                presentationProperties: {
                  published: true
                },
                fields: 'published'
              }
            }]
          };
          
          const slidesResponse = UrlFetchApp.fetch(slidesUrl, {
            method: 'POST',
            headers: {
              'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
              'Content-Type': 'application/json'
            },
            payload: JSON.stringify(batchUpdatePayload)
          });
          
          if (slidesResponse.getResponseCode() === 200) {
            console.log(`  ✅ Published via Slides API`);
          } else {
            console.log(`  ⚠️ Slides API response: ${slidesResponse.getResponseCode()}`);
          }
        } catch (slidesError) {
          console.log(`  ⚠️ Slides API error: ${slidesError.message}`);
        }
        
        // Generate published URL
        const publishedUrl = `https://docs.google.com/presentation/d/e/2PACX-${presentationId}/pub?start=false&loop=false&delayms=3000`;
        
        results.push({
          combination: i + 1,
          presentationId: presentationId,
          editUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
          publishedUrl: publishedUrl,
          success: true
        });
        
        console.log(`  🔗 Edit URL: https://docs.google.com/presentation/d/${presentationId}/edit`);
        console.log(`  🌐 Published URL: ${publishedUrl}`);
        
      } catch (error) {
        console.error(`  ❌ Failed to publish presentation ${i + 1}: ${error.message}`);
        results.push({
          combination: i + 1,
          presentationId: presentationId,
          success: false,
          error: error.toString()
        });
      }
      
      // Small delay between requests
      Utilities.sleep(1000);
    }
    
    console.log('\n✅ Publishing complete!');
    console.log('📊 Results summary:');
    results.forEach(result => {
      if (result.success) {
        console.log(`   ${result.combination}: ✅ Published`);
      } else {
        console.log(`   ${result.combination}: ❌ Failed - ${result.error}`);
      }
    });
    
    return {
      success: true,
      totalProcessed: results.length,
      successfulPublications: results.filter(r => r.success).length,
      failedPublications: results.filter(r => !r.success).length,
      results: results,
      message: 'Acid test presentations published for screenshot access'
    };
    
  } catch (error) {
    console.error('💥 Publishing process failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

// Alternative: Publish individual presentation
function publishPresentation(presentationId) {
  console.log(`📤 Publishing single presentation: ${presentationId}`);
  
  try {
    // Make publicly viewable
    const driveUrl = `https://www.googleapis.com/drive/v3/files/${presentationId}/permissions`;
    
    const response = UrlFetchApp.fetch(driveUrl, {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        role: 'reader',
        type: 'anyone'
      })
    });
    
    if (response.getResponseCode() === 200) {
      console.log('✅ Successfully published presentation');
      return {
        success: true,
        presentationId: presentationId,
        editUrl: `https://docs.google.com/presentation/d/${presentationId}/edit`,
        publishedUrl: `https://docs.google.com/presentation/d/e/2PACX-${presentationId}/pub?start=false&loop=false&delayms=3000`
      };
    } else {
      throw new Error(`Drive API error: ${response.getResponseCode()}`);
    }
    
  } catch (error) {
    console.error('❌ Failed to publish:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}