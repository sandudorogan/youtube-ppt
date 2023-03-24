# YouTube PowerPoint Generator

This script creates a PowerPoint presentation from a YouTube video by extracting keyframes, cropping the speaker, and generating slides from the unique images.

## Features

* Downloads a YouTube video based on the URL provided.
* Allows specifying a crop rectangle to focus on the speaker.
* Supports start and end times for the presentation.
* Removes duplicate slides based on Mean Squared Error (MSE) threshold.
* Caches downloaded videos and generated images for faster processing.
* Allows specifying the output path for the PowerPoint file.

## Requirements

* Python 3.6 or higher
* Install the required packages using `pip install -r requirements.txt`

## Usage

```sh
python youtube_ppt.py <url> [--crop x,y,width,height] [--start MM:SS] [--end MM:SS] [--output output_path]
```

* `<url>`: YouTube video URL (required).
* `--crop`: Crop rectangle in format x,y,width,height (optional).
* `--start`: Start time of the presentation in format MM:SS (optional, default is "00:00").
* `--end`: End time of the presentation in format MM:SS (optional, default is the end of the video).
* `--output`: Output PowerPoint file path (optional, default is the current working directory with the video ID as the filename).

## Example

```sh
python youtube_ppt.py "https://www.youtube.com/watch?v=RshscGDLQD8" --crop 200,0,1500,1000 --start 00:36 --end 16:01
```

This command will download the YouTube video at the specified URL, crop the speaker using the provided rectangle, create a PowerPoint presentation starting at 00:36 and ending at 16:01, and save the resulting PowerPoint file in the current working directory with the video ID as the filename.

## Notes

* The video and images are cached to speed up processing for repeated runs. If the crop values change, the script will generate new images and cache them separately.
* The script will attempt to use cached resources when available, but it will download the video and regenerate images if necessary.
* The script uses Mean Squared Error (MSE) to compare images and remove duplicates. The threshold for MSE can be adjusted in the script.
* The script requires an internet connection to download YouTube videos.

## License

This project is released under the MIT License.

