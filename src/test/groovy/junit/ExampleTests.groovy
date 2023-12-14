package junit

import de.ser.doxis4.agentserver.AgentExecutionResult
import eng.ser.com.ImportProjectDocs
import org.junit.*

class ExampleTests {

    Binding binding

    @BeforeClass
    static void initSessionPool() {
        AgentTester.initSessionPool()
    }

    @Before
    void retrieveBinding() {
        binding = AgentTester.retrieveBinding()
    }

    @Test
    void testForAgentResult() {
        //def agent = new ImportEngDocs()
        def agent = new ImportProjectDocs();

        //String ids ="SD06D_QCON24859b3414-ee73-4151-9b0d-0e5b93edf494182023-10-13T14:28:26.483Z011"
        String ids ="SD06D_QCON245c279fe6-4e80-4c6b-ab13-f398abe33124182023-12-12T07:45:50.475Z011"



        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = ids

        def result = (AgentExecutionResult)agent.execute(binding.variables)
        assert result.resultCode == 0
    }

    /*
        def agent = new GroovyAgent()
        binding["AGENT_EVENT_OBJECT_CLIENT_ID"] = "SD0bGENERIC_DOC24d79641cb-c437-4e39-81af-3051824baa25182021-01-14T10:09:04.857Z011"
        def result = (AgentExecutionResult)agent.execute(binding.variables)
        assert result.resultCode == 0
        assert result.executionMessage.contains("Linux")
        assert agent.eventInfObj instanceof IDocument
     */


    @Test
    void testForJavaAgentMethod() {
        //def agent = new JavaAgent()
        //agent.initializeGroovyBlueline(binding.variables)
        //assert agent.getServerVersion().contains("Linux")
    }

    @After
    void releaseBinding() {
        AgentTester.releaseBinding(binding)
    }

    @AfterClass
    static void closeSessionPool() {
        AgentTester.closeSessionPool()
    }
}
