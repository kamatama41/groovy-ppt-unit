import groovy.json.JsonSlurper
import org.junit.internal.runners.statements.InvokeMethod
import org.junit.runners.BlockJUnit4ClassRunner
import org.junit.runners.model.FrameworkMethod
import org.junit.runners.model.InitializationError
import org.junit.runners.model.Statement

/**
 * JSONファイルを読み込んで、テスト対象クラスに
 * インジェクションする{@Runner}実装です。
 */
class GroovyPPTTestRunner extends BlockJUnit4ClassRunner {
    PPTPresentation ppt

    GroovyPPTTestRunner(Class<?> klass) throws InitializationError {
        super(klass)
        ppt = perseJson(klass.name)
    }

    static def perseJson(className) {
        def parser = new JsonSlurper()
        def text = new File(className + '.json').text
        def root = parser.parseText(text)
        new PPTPresentation(root)
    }

    @Override
    Statement methodInvoker(FrameworkMethod method, Object test) {
        return new InnerInvoker(method, test);
    }

    private class InnerInvoker extends InvokeMethod {
        InnerInvoker(FrameworkMethod testMethod, Object target) {
            super(testMethod, target);
            target.presentation = new PPTPresentation(ppt)
        }
        @Override
        public void evaluate() throws Throwable {
            super.evaluate();
        }
    }
}
