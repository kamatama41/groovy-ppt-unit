import org.junit.runner.RunWith
import org.junit.Test

@RunWith(GroovyPPTTestRunner)
class SampleTest {
    PPTPresentation presentation

    @Test
    void スライドの枚数確認() {
        assert presentation.slides.size() == 3
    }

    @Test
    void 表紙の内容確認() {
        assert presentation.slides[0].title == 'GroovyでPPTテスト'
        assert presentation.slides[0].shapes[1].text == '@kamatama_41'
    }

    @Test
    void 目次の確認() {
        assert presentation.slides[1].title == '目次'
        // 改行は削除される
        assert presentation.slides[1].shapes[1].text == 'Power PointでVBA?VBAからGroovyを呼び出すスライドショーの内容を書き出す/読み込む'
    }

    @Test
    void 複数テキストボックスの確認() {
        assert presentation.slides[2].title == '複数テキストボックスがある場合'
        assert presentation.slides[2].shapes[1].text == 'テキスト１'
        assert presentation.slides[2].shapes[2].text == 'テキスト２'
        assert presentation.slides[2].shapes[3].text == 'テキスト３'
    }
}