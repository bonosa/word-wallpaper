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

            // Add click event to set wallpaper as background/header/footer
            img.addEventListener('click', () => displayOptions(photo));
            galleryDiv.appendChild(img);

            // Add attribution
            const attribution = document.createElement('p');
            attribution.innerHTML = `Photo by <a href="${photo.user.links.html}?utm_source=word_addin&utm_medium=referral" target="_blank">${photo.user.name}</a> on <a href="https://unsplash.com/?utm_source=word_addin&utm_medium=referral" target="_blank">Unsplash</a>`;
            galleryDiv.appendChild(attribution);
        });
    }

    // Display options for background/header/footer
    function displayOptions(photo) {
        const option = prompt(
            "Choose where to insert the image:\n1. Background\n2. Header\n3. Footer",
            "1"
        );
        if (option === "1") {
            setDocumentBackground(photo);
        } else if (option === "2") {
            setHeaderImage(photo);
        } else if (option === "3") {
            setFooterImage(photo);
        } else {
            alert("Invalid option selected.");
        }
    }

    // Set image as document background
    async function setDocumentBackground(photo) {
        try {
            await Word.run(async (context) => {
                const sections = context.document.sections;
                sections.load("items");
                await context.sync();

                const body = sections.items[0].body;
                body.insertInlinePictureFromBase64(photo.urls.regular, Word.InsertLocation.replace);
                await context.sync();
            });
            alert("Image set as document background!");
        } catch (error) {
            console.error("Error setting background:", error);
        }
    }

    // Set image in the header
    async function setHeaderImage(photo) {
        try {
            await Word.run(async (context) => {
                const sections = context.document.sections;
                sections.load("items");
                await context.sync();

                const header = sections.items[0].getHeader("primary");
                header.insertInlinePictureFromBase64(photo.urls.regular, Word.InsertLocation.replace);
                await context.sync();
            });
            alert("Image set in the header!");
        } catch (error) {
            console.error("Error setting header image:", error);
        }
    }

    // Set image in the footer
    async function setFooterImage(photo) {
        try {
            await Word.run(async (context) => {
                const sections = context.document.sections;
                sections.load("items");
                await context.sync();

                const footer = sections.items[0].getFooter("primary");
                footer.insertInlinePictureFromBase64(photo.urls.regular, Word.InsertLocation.replace);
                await context.sync();
            });
            alert("Image set in the footer!");
        } catch (error) {
            console.error("Error setting footer image:", error);
        }
    }

    // Load wallpapers when taskpane opens
    const photos = await fetchWallpapers();
    renderGallery(photos);
});
