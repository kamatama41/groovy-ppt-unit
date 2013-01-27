class PPTPresentation {
    def slides = []
    PPTPresentation(root) {
        init(root)
    }
    PPTPresentation(PPTPresentation ppt) {
        init(ppt.slides)
    }

    def init (srcSlides) {
        srcSlides.eachWithIndex { srcSlide, index ->
            slides[index] = new Slide(title: srcSlide.title, text: srcSlide.text)
        }
    }

    class Slide {
        String title
        String text
    }
}