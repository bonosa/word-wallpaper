document.addEventListener('DOMContentLoaded', async () => {
    console.log("DOM fully loaded and parsed.");

    const galleryDiv = document.getElementById('wallpaper-gallery');
    const statusDiv = document.getElementById('status');

    if (!galleryDiv) {
        console.error("Error: 'wallpaper-gallery' div not found!");
        return;
    }
    if (!statusDiv) {
        console.error("Error: 'status' div not found!");
        return;
    }

    console.log("Both DOM elements found. Proceeding...");

    // Fetch wallpapers from Unsplash
    async function fetchWallpapers() {
        try {
            const response = await fetch(
                `https://api.unsplash.com/photos?per_page=10&client_id=4b2pM-iD5Ltty5GOtVnOs7KDOUopxaTdijfXaDHhXcY`
            );
            const photos = await response.json();
            console.log("Unsplash API Response:", photos);
            return photos;
        } catch (error) {
            console.error("Error fetching wallpapers:", error);
            return [];
        }
    }

    // Render wallpapers in the gallery
    function renderGallery(photos) {
        console.log("Rendering photos:", photos);

        photos.forEach(photo => {
            if (!photo || !photo.urls || !photo.urls.small) {
                console.warn("Invalid photo object:", photo); // Log invalid objects
                return; // Skip this entry
            }

            // Create and display the image
            const img = document.createElement('img');
            img.src = photo.urls.small;
            img.alt = photo.description || 'Unsplash Image';
            img.style.cursor = 'pointer';

            // Add click event to set wallpaper
            img.addEventListener('click', () => setWallpaper(photo));
            galleryDiv.appendChild(img);

            // Add attribution
            const attribution = document.createElement('p');
            attribution.innerHTML = `Photo by <a href="${photo.user.links.html}?utm_source=word_addin&utm_medium=referral" target="_blank">${photo.user.name}</a> on <a href="https://unsplash.com/?utm_source=word_addin&utm_medium=referral" target="_blank">Unsplash</a>`;
            galleryDiv.appendChild(attribution);
        });
    }

    // Trigger Unsplash download and set as wallpaper
    async function setWallpaper(photo) {
        try {
            statusDiv.textContent = 'Applying wallpaper...';

            await fetch('/api/download-image', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ downloadUrl: photo.links.download_location }),
            });

            statusDiv.textContent = 'Wallpaper applied successfully!';
        } catch (error) {
            console.error("Error applying wallpaper:", error);
            statusDiv.textContent = 'Failed to apply wallpaper';
        }
    }

    // Load wallpapers when taskpane opens
    const photos = await fetchWallpapers();
    renderGallery(photos);
});
