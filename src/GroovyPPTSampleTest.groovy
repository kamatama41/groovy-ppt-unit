import org.junit.runner.RunWith
import org.junit.Test

@RunWith(GroovyPPTTestRunner.class)
class GroovyPPTSampleTest {
    PPTPresentation presentation

     @Test
     void スライドは2枚である() {
         assert presentation.slides.size() == 2
     }

    @Test
    void 表紙の内容確認() {
        assert presentation.slides[0].title == 'GroovyでPPTテスト'
        assert presentation.slides[0].text == '@kamatama_41'
    }

    @Test
    void 目次の確認() {
        assert presentation.slides[1].title == '目次'
        // 改行は削除される
        assert presentation.slides[1].text == 'Power PointでVBA?aVBAからGroovyを呼び出すスライドショーの内容を書き出す/読み込む'
    }
}
