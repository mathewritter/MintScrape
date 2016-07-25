using System;

namespace MintScrape.Window {
    public static class Size {
        public static void Resize() {
            var currentHeight = Console.WindowHeight;
            var currentWidth = Console.WindowWidth;
            Console.SetWindowSize(currentWidth*2, currentHeight*2);
        }
    }
}