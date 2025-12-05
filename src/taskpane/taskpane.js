/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    
    // Setup modal close handlers
    setupModal();
  }
});

export async function run() {
  return Word.run(async (context) => {
    // Get the selected content
    const selection = context.document.getSelection();
    const images = selection.inlinePictures;
    images.load("items");
    
    await context.sync();
    
    // Check if an image is selected
    if (images.items.length > 0) {
      const image = images.items[0];
      image.load("width, height");
      
      // Get the base64 image data using the method
      const base64Result = image.getBase64ImageSrc();
      
      await context.sync();
      
      // Access the value property
      const base64Data = base64Result.value;
      
      // Display the image in modal
      displayImageModal(base64Data, image.width, image.height);
    } else {
      // Show message if no image is selected
      const container = document.getElementById("image-container");
      if (container) {
        container.innerHTML = "<p style='color: #d13438; padding: 10px;'>Please select an image in the document first, then click the button.</p>";
      }
    }
  });
}

function setupModal() {
  const modal = document.getElementById("modal");
  const closeBtn = document.querySelector(".modal-close");
  
  // Close when clicking the X
  closeBtn.onclick = function() {
    modal.classList.remove("active");
  }
  
  // Close when clicking outside the image
  modal.onclick = function(event) {
    if (event.target === modal) {
      modal.classList.remove("active");
    }
  }
  
  // Close on ESC key
  document.addEventListener("keydown", function(event) {
    if (event.key === "Escape") {
      modal.classList.remove("active");
    }
  });
}

function displayImageModal(base64Data, width, height) {
  // Ensure the base64 data has the proper data URI prefix
  let imageUrl = base64Data;
  if (base64Data && !base64Data.startsWith('data:')) {
    imageUrl = `data:image/png;base64,${base64Data}`;
  }
  
  // Set the image
  const modalImg = document.getElementById("modal-image");
  const modalInfo = document.getElementById("modal-info");
  
  modalImg.src = imageUrl;
  modalInfo.innerHTML = `Original size: ${width} x ${height}px<br><small>Click outside or press ESC to close</small>`;
  
  // Show the modal
  const modal = document.getElementById("modal");
  modal.classList.add("active");
  
  // Update the container to show success
  const container = document.getElementById("image-container");
  if (container) {
    container.innerHTML = "<p style='color: #107c10; padding: 10px;'>✓ Image opened in viewer</p>";
  }
}