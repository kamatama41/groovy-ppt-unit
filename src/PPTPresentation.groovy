class PPTPresentation {
    def slides = []
    PPTPresentation(ppt) {
        init(ppt.slides)
    }

    def init (srcSlides) {
        srcSlides.eachWithIndex { slide, slideIndex ->
            slides[slideIndex] = new Slide()
            slide.shapes.eachWithIndex {  shape, shapeIndex ->
                slides[slideIndex].shapes[shapeIndex] = new Shape(text: shape.text)
            }

        }
    }

    class Slide {
        def shapes = []
        def getTitle () {
            shapes[0].text
        }
    }
    class Shape {
        String text
    }
}